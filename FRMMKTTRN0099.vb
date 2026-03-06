Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Public Class FRMMKTTRN0099
    Inherits System.Windows.Forms.Form
    Private Enum ENUMGRIDDTLS
        ENUM_Mark = 1
        ENUM_TEMP_INVNO
        ENUM_CUSTOMERCODE
        ENUM_CUSTNAME
        ENUM_ITEMCODE
        ENUM_RATEDIFF
        ENUM_QTY
        ENUM_VALUE
        ENUM_TAXABE_AMT
        ENUM_TOTAL_INVAMT
        ENUM_AR_DEBITNOTE
    End Enum
#Region "REVISION HISTORY"
    'Copyright          :   MIND Ltd.
    'Form Name          :   FRMMKTTRN0083
    'Created By         :   Ekta Uniyal
    'Created on         :   12 Jun 2014
    'Description        :   ASN File Generation 
    'Issue ID           :   10613846 — eMPro- CSV file generation for ASN 
    '********************************************************************************************************************
    'Issue ID           :   10613846 — eMPro- CSV File Generation: for more than one invoices, single file will generate now
    'Revised by         :   Prashant Rajpal
#End Region

    Dim mintIndex As Short
    Dim mintFormIndex As Short
    Dim strSql As String = String.Empty
    Dim mblnSameSeries As Boolean
    Dim mSaleConfNo As Double
    Dim mblnpostinfin As Boolean
    Dim mstrMasterString As String
    Dim mstrDetailString As String
    Dim strCustCode As String
    Dim mstrPurposeCode As String
    Dim mstrSuffix As String
    Dim mintnocopies As Short
    Dim mstrReportFilename As String
    Dim mbln_SHIPPING_ADDRESS As Boolean
    Dim frmRpt As eMProCrystalReportViewer
    Dim CR As ReportDocument
    Dim lngCounter As Integer
    Dim mInvNo As String
    Dim mstrInvoiceDate As String
    Dim Ctlinvoice As String
    Dim mstrunlockedfinalmsg As String
    Dim mstlockedfinalmsg As String
    Dim mstrinvoiceno As String
    Dim mstrstrprevinvoice As String
    Dim strcustomercode As Object
    Dim mblnsamelockingdate As Boolean = False
    Dim FLAGISCUSTOMER_PDF_REQ As Boolean = False ''ADDED BY SUMIT KUMAR ON 22 AUG 2019
    Dim mblnEwayFunctionality As Boolean = False
    Dim mblnSUPP_INVOICE_ONEITEM_MULTIPLE As Boolean
    Dim mPdfReaderPath As String = String.Empty
    Dim pdfPrintProcID As Integer = 0


#Region "FORM EVENTS"

    Private Sub FRMMKTTRN0083_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            System.Windows.Forms.Application.DoEvents()
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub Form_Initialize_Renamed()
        On Error GoTo Err_Handler
        gobjDB = New ClsResultSetDB
        gobjDB.GetResult("SELECT postinfin FROM sales_parameter where UNIT_CODE = '" & gstrUnitId & "' ")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        gobjDB.ResultSetClose()
        gobjDB = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FRMMKTTRN0083_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            lblpdfcreate.Visible = False
            btnSavePDFBinary.Visible = False
            Dim rsSalesParameter As New ClsResultSetDB
            mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)
            Call FitToClient(Me, FrmMain, ctlFormHeader, grpBoxBtns, 300)
            dtToDate.Format = DateTimePickerFormat.Custom
            dtToDate.CustomFormat = gstrDateFormat
            dtToDate.Value = GetServerDate()

            dtFromDate.Format = DateTimePickerFormat.Custom
            dtFromDate.CustomFormat = gstrDateFormat
            dtFromDate.Value = GetServerDate()

            txtFromInvoice.Text = ""
            txtToInvoice.Text = ""
            OptGenerate.Checked = True
            optDrnote.Checked = True
            OptOriginalcopy.Checked = True
            Call Form_Initialize_Renamed()
            InitializeSpread_FileDtls()
            rsSalesParameter.GetResult("SELECT REQD_SHIPPING_ADDRESS_SUPPLEMENTARY FROM SALES_PARAMETER where unit_code='" & gstrUnitId & "' ")
            If rsSalesParameter.GetNoRows > 0 Then
                mbln_SHIPPING_ADDRESS = rsSalesParameter.GetValue("REQD_SHIPPING_ADDRESS_SUPPLEMENTARY")
            End If
            rsSalesParameter.ResultSetClose()
            rsSalesParameter = Nothing
            btn_print.Enabled = False
            mblnEwayFunctionality = Convert.ToBoolean(Find_Value("SELECT ISNULL(SUPP_EWAY_BILL_FUNCTIONALITY,0) AS SUPP_EWAY_BILL_FUNCTIONALITY  FROM SALECONF (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='F' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"))
            mblnSUPP_INVOICE_ONEITEM_MULTIPLE = CBool(Find_Value("SELECT SUPP_INVOICE_ONEITEM_MULTIPLE FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'"))
            mPdfReaderPath = (SqlConnectionclass.ExecuteScalar("SELECT PDFREADERPATH FROM GLOBAL_FLAG (NOLOCK)")).ToString()

            '' PRAVEEN DIGITAL SIGN
            mblnISTrueSignRequired = CBool(Find_Value("SELECT ISNULL(IS_TRUE_SIGN_REQUIRED,0) FROM gen_unitmaster (NOLOCK) WHERE Unt_CodeID='" + gstrUNITID + "'"))
            mblnAPIUrl = Find_Value("Select API_Url from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            mblnPFX_ID = Find_Value("Select PFX_ID from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            mblnPFX_Pass = Find_Value("Select PFX_password from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            mblnAPI_Key = Find_Value("Select API_Key from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub FRMMKTTRN0083_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try
            Me.Dispose()
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub FRMMKTTRN0083_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0083_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            If Shift <> 0 Then Exit Sub
            If KeyCode = System.Windows.Forms.Keys.F4 Then Call ctlFormHeader_Click(ctlFormHeader, New System.EventArgs()) : Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "FORM CONTROL EVENTS"






    Private Sub cmdFrmInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFrmInvoice.Click

        Dim strHelp() As String = Nothing

        Try
            If dtFromDate.Value > dtToDate.Value Then
                MessageBox.Show("[From date] should be less than or equal to [To date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtFromDate.Focus()
                Exit Sub
            Else
                If OptGenerate.Checked = True Then
                    If optDrnote.Checked = True Then
                        strSql = " SELECT distinct CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM supplementaryinv_hdr (nolock) " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0  and doc_no <>0 and drcr='DR'" & _
                         " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                         " ORDER BY Invoice_No"
                    Else
                        strSql = " SELECT distinct CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM supplementaryinv_hdr (nolock) " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0  and doc_no <>0 and drcr='CR'" & _
                         " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                         " ORDER BY Invoice_No"
                    End If
                Else
                    If optDrnote.Checked = True Then
                        If mblnEwayFunctionality Then
                            strSql = "SELECT DISTINCT T.Invoice_No,T.Invoice_Date FROM (" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='DR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND S.EWAY_IRN_REQUIRED='N'" & _
                                            " UNION" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='DR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                        " ) AS T " & _
                                        " ORDER BY T.Invoice_No"
                        Else
                            strSql = " SELECT distinct CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                                        " FROM supplementaryinv_hdr " & _
                                        " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                        " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0  and doc_no <>0 and drcr='DR'" & _
                                        " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                                        " ORDER BY Invoice_No"
                        End If
                    Else
                        If mblnEwayFunctionality Then
                            strSql = "SELECT DISTINCT T.Invoice_No,T.Invoice_Date FROM (" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='CR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND S.EWAY_IRN_REQUIRED='N'" & _
                                            " UNION" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='CR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                        " ) AS T " & _
                                        " ORDER BY T.Invoice_No"
                        Else
                            strSql = " SELECT distinct CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                                          " FROM supplementaryinv_hdr " & _
                                          " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                          " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0  and doc_no <>0 and drcr='CR'" & _
                                          " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                                          " ORDER BY Invoice_No"
                        End If
                    End If
                End If


                strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
                If Not (UBound(strHelp) <= 0) Then
                    If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                        MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtFromInvoice.Text = ""
                        Exit Sub
                    Else
                        txtFromInvoice.Text = strHelp(0).Trim
                        txtFromInvoice_Validating(txtFromInvoice, New System.ComponentModel.CancelEventArgs(False))
                    End If
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub cmdToInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToInvoice.Click

        Dim strHelp() As String = Nothing

        Try
            If dtFromDate.Value > dtToDate.Value Then
                MessageBox.Show("[From date] should be less than or equal to [To date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtFromDate.Focus()
                Exit Sub
            Else
                strSql = String.Empty
                If OptGenerate.Checked = True Then
                    If optDrnote.Checked = True Then
                        strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM supplementaryinv_hdr (nolock) " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                         " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                         " AND DRCR= 'DR'" & _
                         " ORDER BY Invoice_No"
                    Else
                        strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM supplementaryinv_hdr (nolock) " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                         " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                         " AND DRCR= 'CR'" & _
                         " ORDER BY Invoice_No"
                    End If
                Else
                    If optDrnote.Checked = True Then
                        If mblnEwayFunctionality Then
                            strSql = "SELECT DISTINCT T.Invoice_No,T.Invoice_Date FROM (" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='DR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND S.EWAY_IRN_REQUIRED='N'" & _
                                            " UNION" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='DR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                        " ) AS T " & _
                                        " ORDER BY T.Invoice_No"
                        Else
                            strSql = " SELECT distinct CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                                        " FROM supplementaryinv_hdr " & _
                                        " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                        " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                                        " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                                        " AND DRCR= 'DR'" & _
                                        " ORDER BY Invoice_No"
                        End If
                       
                    Else
                        If mblnEwayFunctionality Then
                            strSql = "SELECT DISTINCT T.Invoice_No,T.Invoice_Date FROM (" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='CR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND S.EWAY_IRN_REQUIRED='N'" & _
                                            " UNION" & _
                                        " SELECT distinct CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date" & _
                                            " FROM supplementaryinv_hdr S " & _
                                            " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                            " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                            " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no <>0 and S.drcr='CR'" & _
                                            " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                            " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                        " ) AS T " & _
                                        " ORDER BY T.Invoice_No"
                        Else
                            strSql = " SELECT distinct CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                                        " FROM supplementaryinv_hdr " & _
                                        " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                        " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                                        " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                                        " AND DRCR= 'CR'" & _
                                        " ORDER BY Invoice_No"
                        End If
                    End If
                End If

                strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
                If Not (UBound(strHelp) <= 0) Then
                    If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                        MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtFromInvoice.Text = ""
                        Exit Sub
                    Else
                        txtToInvoice.Text = strHelp(0).Trim
                        txtToInvoice_Validating(txtToInvoice, New System.ComponentModel.CancelEventArgs(False))
                    End If
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtToInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtToInvoice.GotFocus
        Try
            With txtToInvoice
                .SelectionStart = 0
                .SelectionLength = txtToInvoice.Text.Trim.Length
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtToInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtToInvoice.KeyDown

        Try
            If e.KeyCode = Keys.F1 Then
                Call cmdToInvoice_Click(cmdToInvoice, New System.EventArgs)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtToInvoice_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtToInvoice.Validating

        Try

            txtToInvoice.Text = Replace(txtToInvoice.Text, "'", "")
            If txtToInvoice.Text.Trim.Length = 0 Then Return
            If OptGenerate.Checked = True Then
                If optDrnote.Checked = True Then
                    strSql = " SELECT Doc_No As Invoice_No " & _
                     " FROM supplementaryinv_hdr (nolock) " & _
                      " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                     " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                     " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                     " AND Doc_No='" & txtToInvoice.Text.Trim & "'" & _
                    " AND DRCR='DR'"
                Else
                    strSql = " SELECT Doc_No As Invoice_No " & _
                     " FROM supplementaryinv_hdr (nolock) " & _
                      " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                     " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                     " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                     " AND Doc_No='" & txtToInvoice.Text.Trim & "' " & _
                    " AND DRCR='CR'"
                End If
            Else
                If optDrnote.Checked = True Then
                    If mblnEwayFunctionality Then
                        strSql = "SELECT DISTINCT T.Invoice_No FROM (" & _
                                     " SELECT distinct S.Doc_No As Invoice_No" & _
                                         " FROM supplementaryinv_hdr S " & _
                                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                         " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtToInvoice.Text.Trim & "' and S.drcr='DR'" & _
                                         " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                         " AND S.EWAY_IRN_REQUIRED='N'" & _
                                         " UNION" & _
                                     " SELECT distinct S.Doc_No As Invoice_No" & _
                                         " FROM supplementaryinv_hdr S " & _
                                         " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                         " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtToInvoice.Text.Trim & "' and S.drcr='DR'" & _
                                         " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                         " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                     " ) AS T "
                    Else
                        strSql = " SELECT Doc_No As Invoice_No " & _
                    " FROM supplementaryinv_hdr " & _
                    " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                    " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                    " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                    " AND Doc_No='" & txtToInvoice.Text.Trim & "' " & _
                    " AND DRCR='DR'"
                    End If
                   
                Else
                    If mblnEwayFunctionality Then
                        strSql = "SELECT DISTINCT T.Invoice_No FROM (" & _
                                     " SELECT distinct S.Doc_No As Invoice_No" & _
                                         " FROM supplementaryinv_hdr S " & _
                                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                         " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtToInvoice.Text.Trim & "' and S.drcr='CR'" & _
                                         " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                         " AND S.EWAY_IRN_REQUIRED='N'" & _
                                         " UNION" & _
                                     " SELECT distinct S.Doc_No As Invoice_No" & _
                                         " FROM supplementaryinv_hdr S " & _
                                         " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                         " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtToInvoice.Text.Trim & "' and S.drcr='CR'" & _
                                         " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                         " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                     " ) AS T "
                    Else
                        strSql = " SELECT Doc_No As Invoice_No " & _
                    " FROM supplementaryinv_hdr " & _
                    " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                    " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                    " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                    " AND Doc_No='" & txtToInvoice.Text.Trim & "' " & _
                   "  AND DRCR='CR'"
                    End If
                End If
            End If
            If IsRecordExists(strSql) = False Then
                MessageBox.Show("Selected Invoice No. does not exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtToInvoice.Text = ""
                txtToInvoice.Focus()
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtFromInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFromInvoice.GotFocus
        Try
            With txtFromInvoice
                .SelectionStart = 0
                .SelectionLength = txtFromInvoice.Text.Trim.Length
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtFromInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFromInvoice.KeyDown

        Try

            If e.KeyCode = Keys.F1 Then
                Call cmdFrmInvoice_Click(cmdFrmInvoice, New System.EventArgs)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtFromInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFromInvoice.KeyPress

        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case 13

            End Select

            KeyAscii = validateKey(txtFromInvoice.Text, Len(Me.txtFromInvoice.Text), KeyAscii, 12, 0)
            e.KeyChar = Chr(KeyAscii)

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub txtFromInvoice_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFromInvoice.Validating

        Try

            txtFromInvoice.Text = Replace(txtFromInvoice.Text, "'", "")
            If txtFromInvoice.Text.Trim.Length = 0 Then Return
            If OptGenerate.Checked = True Then

                If optDrnote.Checked = True Then
                    strSql = " SELECT Doc_No As Invoice_No " & _
                     " FROM supplementaryinv_hdr (nolock)" & _
                      " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                     " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                     " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                     " AND Doc_No='" & txtFromInvoice.Text.Trim & "' " & _
                    " AND DRCR= 'DR'"
                Else
                    strSql = " SELECT Doc_No As Invoice_No " & _
                     " FROM supplementaryinv_hdr (nolock)" & _
                      " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                     " AND Bill_Flag = 0 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                     " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                     " AND Doc_No='" & txtFromInvoice.Text.Trim & "' " & _
                    " AND DRCR= 'CR'"
                End If

            Else

                If optDrnote.Checked = True Then
                    If mblnEwayFunctionality Then
                        strSql = "SELECT DISTINCT T.Invoice_No FROM (" & _
                                       " SELECT distinct S.Doc_No As Invoice_No" & _
                                           " FROM supplementaryinv_hdr S " & _
                                           " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                           " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtFromInvoice.Text.Trim & "' and S.drcr='DR'" & _
                                           " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                           " AND S.EWAY_IRN_REQUIRED='N'" & _
                                           " UNION" & _
                                       " SELECT distinct S.Doc_No As Invoice_No" & _
                                           " FROM supplementaryinv_hdr S " & _
                                           " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                           " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                           " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtFromInvoice.Text.Trim & "' and S.drcr='DR'" & _
                                           " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                           " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                       " ) AS T "
                    Else
                        strSql = " SELECT Doc_No As Invoice_No " & _
                     " FROM supplementaryinv_hdr (nolock)" & _
                      " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                     " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                     " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                     " AND Doc_No='" & txtFromInvoice.Text.Trim & "' " & _
                    " AND DRCR= 'DR'"
                    End If
                Else
                    If mblnEwayFunctionality Then
                        strSql = "SELECT DISTINCT T.Invoice_No FROM (" & _
                                     " SELECT distinct S.Doc_No As Invoice_No" & _
                                         " FROM supplementaryinv_hdr S " & _
                                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                         " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtFromInvoice.Text.Trim & "' and S.drcr='CR'" & _
                                         " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                         " AND S.EWAY_IRN_REQUIRED='N'" & _
                                         " UNION" & _
                                     " SELECT distinct S.Doc_No As Invoice_No" & _
                                         " FROM supplementaryinv_hdr S " & _
                                         " LEFT JOIN Supplementary_IRN I ON S.UNIT_CODE=I.UNIT_CODE AND S.VOUCHER_NO=I.VO_NO " & _
                                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                                         " AND S.Bill_Flag = 1 And ISNULL(S.CANCEL_FLAG, 0) = 0  and S.doc_no ='" & txtFromInvoice.Text.Trim & "' and S.drcr='CR'" & _
                                         " AND S.UNIT_CODE = '" & gstrUnitId & "'" & _
                                         " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(I.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(I.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                                     " ) AS T "
                    Else
                        strSql = " SELECT Doc_No As Invoice_No " & _
                         " FROM supplementaryinv_hdr (nolock)" & _
                          " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                         " AND UNIT_CODE = '" & gstrUnitId & "'" & _
                         " AND Doc_No='" & txtFromInvoice.Text.Trim & "' " & _
                        " AND DRCR= 'CR'"
                    End If
                End If

            End If

            If IsRecordExists(strSql) = False Then
                MessageBox.Show("Selected Invoice No. does not exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtFromInvoice.Text = ""
                txtFromInvoice.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtFromInvoice_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFromInvoice.TextChanged

        Try
            If txtFromInvoice.Text.Length = 0 Then
                txtFromInvoice.Text = ""
                txtToInvoice.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtToInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtToInvoice.KeyPress

        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case 13

            End Select

            KeyAscii = validateKey(txtFromInvoice.Text, Len(Me.txtFromInvoice.Text), KeyAscii, 12, 0)
            e.KeyChar = Chr(KeyAscii)



        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub txtToInvoice_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtToInvoice.TextChanged

        Try
            If txtToInvoice.Text.Length = 0 Then
                txtFromInvoice.Text = ""
                txtToInvoice.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub ctlFormHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader.Click

        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub dtFromDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtFromDate.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case 13
                    dtToDate.Focus()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dtToDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtToDate.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case 13
                    txtFromInvoice.Focus()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "ROUTINES"

    Private Sub GenerateASNFile()

        Dim oSqlDr As SqlDataReader = Nothing
        Dim Dt As New DataTable
        Dim sw As StreamWriter = Nothing
        Dim iColCount As Integer = 0
        Dim intCol As Integer = 0
        Dim intLoopCounter As Integer = 0
        Dim strFinalQry As String = String.Empty
        Dim strValDocNo As String = String.Empty
        Dim strQuery As String = String.Empty
        Dim strFileLocation As String = String.Empty
        Dim strASNFilePath As String = String.Empty
        Dim fs As FileStream = Nothing
        Dim Obj_FSO As Scripting.FileSystemObject = Nothing

        Try
            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
            If Dt.Rows.Count > 0 Then Dt.Dispose()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Not sw Is Nothing Then sw.Close()
            If Not fs Is Nothing Then fs.Close()
            If Not Obj_FSO Is Nothing Then Obj_FSO = Nothing
            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
            If Dt.Rows.Count > 0 Then Dt.Dispose()
        End Try

    End Sub
    Private Function IsGSTINSAME(ByVal strCustomerCode As String) As Boolean
        If SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_Id,'') GSTIN_Id From customer_vendor_vw (NOLOCK) Where UNIT_CODE='" & gstrUNITID & "' And Customer_Code='" & strCustomerCode & "'") = SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_ID,'') GSTIN_ID From Gen_UnitMaster Where Unt_CodeId='" & gstrUNITID & "'") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function Validate_Data() As Boolean

        Dim strValidate As String = String.Empty

        Try


            If txtFromInvoice.Text.Trim.Length <= 0 Then
                strValidate = strValidate + Chr(13) & "From Invoice cannot be Blank."
                txtFromInvoice.Focus()
            End If

            If txtToInvoice.Text.Trim.Length <= 0 Then
                strValidate = strValidate + Chr(13) & "To Invoice cannot be Blank."
                txtToInvoice.Focus()
            End If

            If dtFromDate.Value > dtToDate.Value Then
                strValidate = strValidate + Chr(13) & "[From date] should be less than or equal to [To date]."
                dtFromDate.Focus()
            End If

            If Val(txtFromInvoice.Text) > Val(txtToInvoice.Text) Then
                strValidate = strValidate + Chr(13) & "[From Invoice] should be less than or equal to [To Invoice]."
                txtFromInvoice.Focus()
            End If

            If strValidate <> String.Empty Then
                Validate_Data = False
                MsgBox(strValidate, MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            Else
                Validate_Data = True
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Function

    Private Function FN_Get_Folder_Path() As String

        Dim strFilePath As String = String.Empty
        Dim strReturnValue As String = String.Empty

        Try
            strSql = " SELECT ISNULL(ASN_HMIL_FilePath,'')as ASN_HMIL_FilePath" & _
                     " FROM Sales_Parameter" & _
                     " WHERE UNIT_CODE = '" & gstrUnitId & "'"
            strFilePath = SqlConnectionclass.ExecuteScalar(strSql)
            If strFilePath <> String.Empty Then
                strReturnValue = strFilePath
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

        FN_Get_Folder_Path = strReturnValue

    End Function

#End Region

#Region "COMMAND BUTTONS"

    Private Sub btnGenerateFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            GenerateASNFile()
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

#End Region

    Private Sub FrmMain_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMain.Enter

    End Sub

    Private Sub BtnFetchshow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnshow.Click

        Try
            btnSavePDFBinary.Visible = False
            If txtFromInvoice.Text = "" Then
                MsgBox("Please select invoice no (From Invoice No) !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtFromInvoice.Focus()
                Exit Sub
            End If
            If txtToInvoice.Text = "" Then
                MsgBox("Please Select invoice no (To Invoice No ) !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtToInvoice.Focus()
                Exit Sub
            End If
            If dtFromDate.Value > dtToDate.Value Then
                MsgBox("Invalid date range !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Exit Sub
            End If
            GETDATA()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
    Private Sub GETDATA()
        Dim sqlCmd As New SqlCommand()
        Dim SqlAdp As New SqlDataAdapter
        Dim DSFILEDTL As New DataSet
        Dim strsql As String = String.Empty
        Dim intLoopCounter As Int32 = 0
        FLAGISCUSTOMER_PDF_REQ = False
        Dim STRDRCR As String
        Try

            With sqlCmd
                .CommandText = "USP_MULTI_SUPP_INV_DETAILS"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                If OptGenerate.Checked = True Then
                    .Parameters.AddWithValue("@MODE", "G")
                End If
                If OptReprint.Checked = True Then
                    .Parameters.AddWithValue("@MODE", "R")
                End If
                If optDrnote.Checked = True Then
                    STRDRCR = "DR"
                Else
                    STRDRCR = "CR"
                End If
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@FROM_DATE", getDateForDB(dtFromDate.Value))
                .Parameters.AddWithValue("@TO_DATE", getDateForDB(dtToDate.Value))
                .Parameters.AddWithValue("@FROM_INVOICE", txtFromInvoice.Text.Trim)
                .Parameters.AddWithValue("@TO_INVOICE", txtToInvoice.Text.Trim)
                .Parameters.AddWithValue("@DRCR", STRDRCR)
                .Parameters.AddWithValue("@EWAY_BILL_FUNCTIONALITY", IIf(mblnEwayFunctionality = True, 1, 0))
                '.Parameters.AddWithValue("@ERR" 
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSFILEDTL)
                .Dispose()
            End With

            If DSFILEDTL.Tables.Count > 0 Then
                'DTL DATA
                If DSFILEDTL.Tables(0).Rows.Count > 0 Then
                    InitializeSpread_FileDtls()
                    With Me.fsprDtls
                        For intLoopCounter = 0 To DSFILEDTL.Tables(0).Rows.Count - 1
                            AddRow()
                            .SetText(ENUMGRIDDTLS.ENUM_TEMP_INVNO, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("TEMP_INVNO").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_CUSTOMERCODE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("CUSTOMER_CODE").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_CUSTNAME, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("CUST_NAME").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_ITEMCODE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("ITEM_CODE").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_RATEDIFF, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("RATE_DIFF").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_QTY, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("QUANTITY").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_TAXABE_AMT, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("TAXABLE_AMT").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_VALUE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("AMOUNT").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_TOTAL_INVAMT, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("TOTAL_AMOUNT").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_AR_DEBITNOTE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("VOUCHER_NO").ToString.Trim)
                        Next
                    End With
                    ''ADDED BY SUMIT KUMAR 22 AUG 2019
                    If DSFILEDTL.Tables(1).Rows.Count > 0 Then
                        If DSFILEDTL.Tables(1).Rows(0).Item(0) > 0 Then
                            FLAGISCUSTOMER_PDF_REQ = True
                        Else
                            FLAGISCUSTOMER_PDF_REQ = False
                        End If
                    End If
                    If OptReprint.Checked Then
                        Dim strUnitCode As String = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT TOP 1 UNIT_CODE FROM INVOICE_PDF_CONFIG (NOLOCK) WHERE INVOICE_TYPE='SUP' AND INVOICE_SUB_TYPE='S' AND UNIT_CODE='" & gstrUnitId & "'"))
                        If FLAGISCUSTOMER_PDF_REQ = True AndAlso Len(strUnitCode) > 0 Then
                            btnSavePDFBinary.Visible = True
                        Else
                            btnSavePDFBinary.Visible = False
                        End If
                    Else
                        btnSavePDFBinary.Visible = False
                    End If
                Else
                    fsprDtls.MaxRows = 0
                    fsprDtls.MaxCols = 0
                    MsgBox("No Data Found For the Date Range Selected!", MsgBoxStyle.Information, "Empro")
                    txtFromInvoice.Text = ""
                    txtToInvoice.Text = ""
                End If
            Else
                MsgBox("No Data Found For the Date Range Selected!", MsgBoxStyle.Information, "Empro")
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            DSFILEDTL.Clear()
            DSFILEDTL.Dispose()
        End Try

    End Sub
    Public Sub AddRow()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

            With fsprDtls
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_Mark : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_CUSTOMERCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_CUSTNAME : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_ITEMCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_RATEDIFF : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_QTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_VALUE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_TAXABE_AMT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_TOTAL_INVAMT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMGRIDDTLS.ENUM_AR_DEBITNOTE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)


            End With

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub InitializeSpread_FileDtls()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.fsprDtls
                .MaxRows = 0
                If OptGenerate.Checked = True Then
                    .MaxCols = ENUMGRIDDTLS.ENUM_AR_DEBITNOTE
                Else
                    .MaxCols = ENUMGRIDDTLS.ENUM_AR_DEBITNOTE
                End If
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_Mark : .Text = "Select" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                If OptGenerate.Checked = True Then
                    .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO : .Text = "Temp. Inv. No" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                Else
                    .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO : .Text = "Inv. No" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                End If
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_CUSTOMERCODE : .Text = "Customer Code" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_CUSTNAME : .Text = "Customer Name" : .set_ColWidth(.Col, 20) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_ITEMCODE : .Text = "Item Code" : .set_ColWidth(.Col, 16) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_RATEDIFF : .Text = "Rate Difference" : .set_ColWidth(.Col, 6) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_QTY : .Text = "QTY" : .set_ColWidth(.Col, 6) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_VALUE : .Text = "Value" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_TAXABE_AMT : .Text = "Taxable Amt." : .set_ColWidth(.Col, 7) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_TOTAL_INVAMT : .Text = "Total Invoice Amt." : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMGRIDDTLS.ENUM_AR_DEBITNOTE : .Text = "AR Debit Note" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)

                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Row2 = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = 1
                .Col2 = .MaxCols
                .Lock = True
                .BlockMode = False
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub OptReprint_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptReprint.CheckedChanged
        Try
            refreshform()
            InitializeSpread_FileDtls()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub refreshform()
        Try
            btnSavePDFBinary.Visible = False
            FLAGISCUSTOMER_PDF_REQ = False
            Me.txtFromInvoice.Text = ""
            Me.txtToInvoice.Text = ""
            Me.fsprDtls.MaxRows = 0
            If OptGenerate.Checked = True Then
                btn_print.Enabled = False
            Else
                btn_print.Enabled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs)
        If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            Close()
            Exit Sub
        End If
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select

    End Sub

    Private Sub btn_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Public Function CreateStringForAccounts(ByVal docno As String) As Boolean
        Dim objRecordSetnew As ADODB.Recordset
        Dim objTmpRecordset As ADODB.Recordset
        Dim rsSalesInvType As ClsResultSetDB
        Dim strRetVal As String
        Dim strRefInvoiceNo As String
        Dim strInvoiceNo As String
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim strCurrencyCode As String
        Dim dblInvoiceAmt As Double
        Dim dblExchangeRate As Double
        Dim dblBasicAmount As Double
        Dim dblBaseCurrencyAmount As Double
        Dim dblTaxAmt As Double
        Dim strTaxType As String
        Dim strCreditTermsID As String
        Dim strBasicDueDate As String
        Dim strPaymentDueDate As String
        Dim strExpectedDueDate As String
        Dim strCustomerGL As String
        Dim strCustomerSL As String
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim strItemGL As String
        Dim strItemSL As String
        Dim strGlGroupId As String
        Dim dblTaxRate As Double
        Dim varTmp As Object
        Dim dblCCShare As Double
        Dim iCtr As Short
        Dim strCustRef As String
        Dim strParamQuery As String
        Dim rsParameterData As ClsResultSetDB
        Dim blnEOU_FLAG As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnInsIncSTax As Boolean
        Dim blnTotalToolCostRoundOff As Boolean
        Dim blnTCSTax As Boolean
        Dim intBasicRoundOffDecimal As Short
        Dim intSaleTaxRoundOffDecimal As Short
        Dim intExciseRoundOffDecimal As Short
        Dim intSSTRoundOffDecimal As Short
        Dim intTCSRoundOffDecimal As Short
        Dim intToolCostRoundOffDecimal As Short
        Dim blnEcssonCVD As Boolean
        Dim intEcssOnCVDRoundOff As Short
        Dim blnECSSTax As Boolean
        Dim intECSRoundOffDecimal As Short
        Dim blnECSSOnSaleTax As Boolean
        Dim intECSSOnSaleRoundOffDecimal As Short
        Dim blnTurnOverTax As Boolean
        Dim intTurnOverTaxRoundOffDecimal As Short
        Dim blnTotalInvoiceAmount As Boolean
        Dim intTotalInvoiceAmountRoundOffDecimal As Short
        Dim dblInvoiceAmtRoundOff_diff As Double
        Dim blnGSTRoundOff As Boolean
        Dim intGSTRoundOffDecimal As Integer
        Dim strTaxCCCode As String


        strTaxCCCode = ""

        On Error GoTo ErrHandler
        objRecordSetnew = Nothing
        objRecordSetnew = New ADODB.Recordset
        objRecordSetnew.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        'objRecordSetnew.Open("SELECT Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,pervalue,Item_Code,Cust_Item_Code,Currency_Code,Rate,Packing,Packing_Amount,Basic_Amount,Accessible_amount,Excise_type,CVD_type,SAD_type,Excise_per,CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,Surcharge_salesTaxType,Surcharge_SalesTax_Per,Surcharge_Sales_Tax_Amount,total_amount,dataPosted,Transport_Type,Vehicle_No,Carriage_Name,SRVDINO,SRVLocation,RejectionPosting,SuppInv_Remarks,remarks,ECESS_Type,ECESS_Per,ECESS_Amount,sales_Quantity,supp_invdetail,TotalInvoiceAmtRoundOff_diff,SECESS_Type,SECESS_Per,SECESS_Amount,MRP,ADDVAT_TYPE,ADDVAT_PER,ADDVAT_AMOUNT  ,AED_TYPE,AED_PER,AED_AMOUNT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,UTGSTTXRT_TYPE,UTGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,CGST_AMT,SGST_AMT,UTGST_AMT,IGST_AMT,CCESS_AMT FROM SupplementaryInv_hdr WHERE Unit_code='" & gstrUnitId & "' and Doc_No='" & invoiceno & "' and Location_Code='" & gstrUNITID  & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        'objRecordSetnew.Open("SELECT Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,pervalue,Item_Code,Cust_Item_Code,Currency_Code,Rate,Packing,Packing_Amount,Basic_Amount,Accessible_amount,Excise_type,CVD_type,SAD_type,Excise_per,CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,Surcharge_salesTaxType,Surcharge_SalesTax_Per,Surcharge_Sales_Tax_Amount,total_amount,dataPosted,Transport_Type,Vehicle_No,Carriage_Name,SRVDINO,SRVLocation,RejectionPosting,SuppInv_Remarks,remarks,ECESS_Type,ECESS_Per,ECESS_Amount,sales_Quantity,supp_invdetail,TotalInvoiceAmtRoundOff_diff,SECESS_Type,SECESS_Per,SECESS_Amount,MRP,ADDVAT_TYPE,ADDVAT_PER,ADDVAT_AMOUNT  ,AED_TYPE,AED_PER,AED_AMOUNT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,UTGSTTXRT_TYPE,UTGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,CGST_AMT,SGST_AMT,UTGST_AMT,IGST_AMT,CCESS_AMT FROM SupplementaryInv_hdr WHERE Unit_code='" & gstrUNITID & "' and Doc_No='" & Ctlinvoice & "' and Location_Code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        'objRecordSetnew.Open("SELECT Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,pervalue,Item_Code,Cust_Item_Code,Currency_Code,Rate,Packing,Packing_Amount,Basic_Amount,Accessible_amount,Excise_type,CVD_type,SAD_type,Excise_per,CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,Surcharge_salesTaxType,Surcharge_SalesTax_Per,Surcharge_Sales_Tax_Amount,total_amount,dataPosted,Transport_Type,Vehicle_No,Carriage_Name,SRVDINO,SRVLocation,RejectionPosting,SuppInv_Remarks,remarks,ECESS_Type,ECESS_Per,ECESS_Amount,sales_Quantity,supp_invdetail,TotalInvoiceAmtRoundOff_diff,SECESS_Type,SECESS_Per,SECESS_Amount,MRP,ADDVAT_TYPE,ADDVAT_PER,ADDVAT_AMOUNT  ,AED_TYPE,AED_PER,AED_AMOUNT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,UTGSTTXRT_TYPE,UTGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,CGST_AMT=(select sum(DIFF_CGST_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),SGST_AMT=(select sum(DIFF_SGST_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),UTGST_AMT,IGST_AMT=(select sum(DIFF_igst_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),CCESS_AMT FROM SupplementaryInv_hdr sh WHERE Unit_code='" & gstrUNITID & "' and Doc_No='" & Ctlinvoice & "' and Location_Code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        objRecordSetnew.Open("SELECT Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,pervalue,Item_Code,Cust_Item_Code,Currency_Code,Rate,Packing,Packing_Amount,Basic_Amount,Accessible_amount,Excise_type,CVD_type,SAD_type,Excise_per,CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,Surcharge_salesTaxType,Surcharge_SalesTax_Per,Surcharge_Sales_Tax_Amount,total_amount,dataPosted,Transport_Type,Vehicle_No,Carriage_Name,SRVDINO,SRVLocation,RejectionPosting,SuppInv_Remarks,remarks,ECESS_Type,ECESS_Per,ECESS_Amount,sales_Quantity,supp_invdetail,TotalInvoiceAmtRoundOff_diff,SECESS_Type,SECESS_Per,SECESS_Amount,MRP,ADDVAT_TYPE,ADDVAT_PER,ADDVAT_AMOUNT  ,AED_TYPE,AED_PER,AED_AMOUNT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,UTGSTTXRT_TYPE,UTGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,CGST_AMT=(select sum(DIFF_CGST_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),SGST_AMT=(select sum(DIFF_SGST_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),UTGST_AMT,IGST_AMT=(select sum(DIFF_igst_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),CCESS_AMT, ISNULL(TCSTax_Type,'')TCSTax_Type,ISNULL(TCSTax_Per,0)TCSTax_Per,ISNULL(TCSTaxAmount,0)TCSTaxAmount FROM SupplementaryInv_hdr sh WHERE Unit_code='" & gstrUNITID & "' and Doc_No='" & Ctlinvoice & "' and Location_Code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSetnew.EOF Then
            MsgBox("Supplementary Invoice details not found", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSetnew.Close()
                objRecordSetnew = Nothing
            End If
            Exit Function
        End If
        strInvoiceNo = CStr(mInvNo)
        If mblnsamelockingdate = True Then
            mstrInvoiceDate = VB6.Format(GetServerDateTime(), "dd-MMM-yyyy")
        Else
            mstrInvoiceDate = VB6.Format(objRecordSetnew.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        End If

        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSetnew.Fields("Currency_Code").Value), "", objRecordSetnew.Fields("Currency_Code").Value))
        dblInvoiceAmt = IIf(IsDBNull(objRecordSetnew.Fields("total_amount").Value), 0, objRecordSetnew.Fields("total_amount").Value)
        dblInvoiceAmtRoundOff_diff = IIf(IsDBNull(objRecordSetnew.Fields("TotalInvoiceAmtRoundOff_diff").Value), 0, objRecordSetnew.Fields("TotalInvoiceAmtRoundOff_diff").Value)
        dblExchangeRate = 1
        strCustCode = Trim(objRecordSetnew.Fields("Account_Code").Value)
        strCustRef = Trim(IIf(IsDBNull(objRecordSetnew.Fields("cust_ref").Value), "", objRecordSetnew.Fields("cust_ref").Value))
        'To Get Refrance Invoice No

        rsSalesInvType = Nothing
        rsSalesInvType = New ClsResultSetDB

        rsSalesInvType.GetResult("Select top 1 1 from SupplementaryInv_hdr where Unit_code='" & gstrUnitId & "' and Doc_no = '" & docno & "' and Location_code = '" & gstrUnitId & "'")
        If rsSalesInvType.GetNoRows > 0 Then
        Else
            rsSalesInvType.ResultSetClose()
            rsSalesInvType = Nothing
            rsSalesInvType = New ClsResultSetDB
            rsSalesInvType.GetResult("Select RefDoc_no from SupplementaryInv_dtl where Unit_code='" & gstrUnitId & "' and Doc_no = '" & docno & "' and Location_code = '" & gstrUnitId & "'")
            If rsSalesInvType.GetNoRows > 0 Then
                rsSalesInvType.MoveFirst()
                strRefInvoiceNo = rsSalesInvType.GetValue("RefDoc_no")
            Else
                Exit Function
            End If
        End If
        rsSalesInvType.ResultSetClose()
        rsSalesInvType = Nothing
        'Retreiving the customer gl, sl and credit term id
        'objTmpRecordset.CursorType.adOpenKeyset()

        objTmpRecordset = Nothing
        objTmpRecordset = New ADODB.Recordset
        objTmpRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Unit_code='" & gstrUnitId & "' and Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objTmpRecordset.EOF Then
            MsgBox("Customer details not found", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSetnew.Close()
                objRecordSetnew = Nothing
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                objTmpRecordset.Close()
                objTmpRecordset = Nothing
            End If
            Exit Function
        End If
        strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_ArCode").Value), "", objTmpRecordset.Fields("Cst_ArCode").Value))
        strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_slCode").Value), "", objTmpRecordset.Fields("Cst_slCode").Value))
        strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
        If strCreditTermsID = "" Then
            MsgBox("Credit Terms not found", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSetnew.Close()
                objRecordSetnew = Nothing
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                objTmpRecordset.Close()
                objTmpRecordset = Nothing
            End If
            Exit Function
        End If
        Dim objCreditTerms As New prj_CreditTerm.clsCR_Term_Resolver
        strRetVal = objCreditTerms.RetCR_Term_Dates("", "INV", strCreditTermsID, mstrInvoiceDate, gstrUnitId, "", "", gstrCONNECTIONSTRING)
        If CheckString(strRetVal) = "Y" Then
            strRetVal = Mid(strRetVal, 3)
            varTmp = Split(strRetVal, "»")
            strBasicDueDate = VB6.Format(varTmp(0), "dd-MMM-yyyy")
            strPaymentDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
            strExpectedDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
        Else
            MsgBox(CheckString(strRetVal), MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSetnew.Close()
                objRecordSetnew = Nothing
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                objTmpRecordset.Close()
                objTmpRecordset = Nothing
            End If
            Exit Function
        End If
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
        strParamQuery = strParamQuery & "ECESSOnCVD_Roundoff=isnull(ECESSOnCVD_Roundoff,0),ECESSOnCVDRoundOff_Decimal= isnull(ECESSOnCVDRoundOff_Decimal,0),"
        strParamQuery = strParamQuery & " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal, ECESSOnSaleTax_Roundoff, ECESSOnSaleTaxRoundOff_Decimal, "
        strParamQuery = strParamQuery & " TurnOverTax_RoundOff, TurnOverTaxRoundOff_Decimal, TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,ISNULL(GSTTAX_ROUNDOFF_DECIMAL,0) GSTTAX_ROUNDOFF_DECIMAL,ISNULL(GSTTAX_ROUNDOFF,0) GSTTAX_ROUNDOFF FROM Sales_Parameter where Unit_code='" & gstrUnitId & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnInsIncSTax = rsParameterData.GetValue("InsInc_SalesTax")
            blnTotalToolCostRoundOff = rsParameterData.GetValue("TotalToolCostRoundoff")
            blnTCSTax = rsParameterData.GetValue("TCSTax_Roundoff")
            intBasicRoundOffDecimal = rsParameterData.GetValue("Basic_Roundoff_decimal")
            intSaleTaxRoundOffDecimal = rsParameterData.GetValue("SalesTax_Roundoff_decimal")
            intExciseRoundOffDecimal = rsParameterData.GetValue("Excise_Roundoff_decimal")
            intSSTRoundOffDecimal = rsParameterData.GetValue("SST_Roundoff_decimal")
            intTCSRoundOffDecimal = rsParameterData.GetValue("TCSTax_Roundoff_decimal")
            intToolCostRoundOffDecimal = rsParameterData.GetValue("TotalToolCostRoundoff_decimal")
            If blnEOU_FLAG = True Then
                blnEcssonCVD = rsParameterData.GetValue("ECESSOnCVD_Roundoff")
                intEcssOnCVDRoundOff = rsParameterData.GetValue("ECESSOnCVDRoundOff_Decimal")
            End If
            blnECSSTax = rsParameterData.GetValue("ECESS_Roundoff")
            intECSRoundOffDecimal = rsParameterData.GetValue("ECESSRoundoff_Decimal")
            blnECSSOnSaleTax = rsParameterData.GetValue("ECESSOnSaleTax_Roundoff")
            intECSSOnSaleRoundOffDecimal = rsParameterData.GetValue("ECESSOnSaleTaxRoundOff_Decimal")
            blnTurnOverTax = rsParameterData.GetValue("TurnOverTax_RoundOff")
            intTurnOverTaxRoundOffDecimal = rsParameterData.GetValue("TurnOverTaxRoundOff_Decimal")
            blnTotalInvoiceAmount = rsParameterData.GetValue("TotalInvoiceAmount_RoundOff")
            intTotalInvoiceAmountRoundOffDecimal = rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal")
            blnGSTRoundOff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
            intGSTRoundOffDecimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Critical, ResolveResString(100))
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Function
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        dblInvoiceAmt = System.Math.Round(dblInvoiceAmt * dblExchangeRate, intTotalInvoiceAmountRoundOffDecimal)

        mstrMasterString = ""
        mstrDetailString = ""
        If gblnGSTUnit = False Then
            If gstrUNITID = "STH" Then
                If optDrnote.Checked = True Then
                    'mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
                    mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
                Else
                    'mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»CR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
                    mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»CR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
                End If
            Else
                If optDrnote.Checked = True Then
                    mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & mstrInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & mstrInvoiceDate & "»"
                    mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
                    mstrMasterString = mstrMasterString & dblInvoiceAmt & "»" & dblInvoiceAmt & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
                Else
                    mstrMasterString = "I»" & strInvoiceNo & "»Cr»»" & mstrInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & mstrInvoiceDate & "»"
                    mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
                    mstrMasterString = mstrMasterString & dblInvoiceAmt & "»" & dblInvoiceAmt & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
                End If
            End If
            'If optDrnote.Checked = True Then
            '    mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & mstrInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & mstrInvoiceDate & "»"
            '    mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            '    mstrMasterString = mstrMasterString & dblInvoiceAmt & "»" & dblInvoiceAmt & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            'Else
            '    mstrMasterString = "I»" & strInvoiceNo & "»Cr»»" & mstrInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & mstrInvoiceDate & "»"
            '    mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            '    mstrMasterString = mstrMasterString & dblInvoiceAmt & "»" & dblInvoiceAmt & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            'End If

        Else
            If optDrnote.Checked = True Then
                'mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
                mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUnitId & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
            Else
                'mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»CR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
                mstrMasterString = "M»123»" & mstrInvoiceDate & "»0»»" & gstrUnitId & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & mstrInvoiceDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»CR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & mstrInvoiceDate & "»0¦"
            End If

        End If
        iCtr = 1
        ''CST/LST/SRT/VAT Posting

        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        objTmpRecordset = Nothing
        objTmpRecordset = New ADODB.Recordset
        objTmpRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient


        If Trim(IIf(IsDBNull(objRecordSetnew.Fields("SalesTax_Type").Value), "", objRecordSetnew.Fields("SalesTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUnitId & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSetnew.Fields("SalesTax_Type").Value), "", objRecordSetnew.Fields("SalesTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
            If strTaxType = "LST" Or strTaxType = "CST" Or strTaxType = "SRT" Or strTaxType = "VAT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("Sales_Tax_Amount").Value), 0, objRecordSetnew.Fields("Sales_Tax_Amount").Value)
                If blnISSalesTaxRoundOff = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intSaleTaxRoundOffDecimal)
                ElseIf blnISSalesTaxRoundOff = True Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("SalesTax_Per").Value), 0, objRecordSetnew.Fields("SalesTax_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSetnew.Close()
                            objRecordSetnew = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '********* FOR ADITIONAL VAT POSTING *********************
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        objTmpRecordset = Nothing
        objTmpRecordset = New ADODB.Recordset
        objTmpRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        If Trim(IIf(IsDBNull(objRecordSetnew.Fields("ADDVAT_TYPE").Value), "", objRecordSetnew.Fields("ADDVAT_TYPE").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUnitId & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSetnew.Fields("ADDVAT_TYPE").Value), "", objRecordSetnew.Fields("ADDVAT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            'If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing 'AMIT
            If strTaxType = "ADVAT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("ADDVAT_AMOUNT").Value), 0, objRecordSetnew.Fields("ADDVAT_AMOUNT").Value)
                If blnISSalesTaxRoundOff = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intSaleTaxRoundOffDecimal)
                ElseIf blnISSalesTaxRoundOff = True Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("ADDVAT_PER").Value), 0, objRecordSetnew.Fields("ADDVAT_PER").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSetnew.Close()
                            objRecordSetnew = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '**********************END OF ADITIONAL VAT POSTING *********************
        'ECS Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        objTmpRecordset = Nothing
        objTmpRecordset = New ADODB.Recordset
        objTmpRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient


        If Trim(IIf(IsDBNull(objRecordSetnew.Fields("ECESS_Type").Value), "", objRecordSetnew.Fields("ECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUnitId & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSetnew.Fields("ECESS_Type").Value), "", objRecordSetnew.Fields("ECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            'If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing 'AMIT

            If strTaxType = "ECS" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("ECESS_Amount").Value), 0, objRecordSetnew.Fields("ECESS_Amount").Value)
                If blnECSSTax = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intECSRoundOffDecimal)
                Else
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("ECESS_Per").Value), 0, objRecordSetnew.Fields("ECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSetnew.Close()
                            objRecordSetnew = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        ''SECESS Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        objTmpRecordset = Nothing
        objTmpRecordset = New ADODB.Recordset
        objTmpRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient


        If Trim(IIf(IsDBNull(objRecordSetnew.Fields("SECESS_Type").Value), "", objRecordSetnew.Fields("SECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUnitId & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSetnew.Fields("SECESS_Type").Value), "", objRecordSetnew.Fields("SECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            'If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing "AMIT
            If strTaxType = "ECSSH" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("SECESS_Amount").Value), 0, objRecordSetnew.Fields("SECESS_Amount").Value)
                If blnECSSTax = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intECSRoundOffDecimal)
                Else
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("SECESS_Per").Value), 0, objRecordSetnew.Fields("SECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSetnew.Close()
                            objRecordSetnew = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If

        '10792712 
        ''AED 
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        objTmpRecordset = Nothing
        objTmpRecordset = New ADODB.Recordset
        objTmpRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        If Trim(IIf(IsDBNull(objRecordSetnew.Fields("AED_TYPE").Value), "", objRecordSetnew.Fields("AED_TYPE").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUnitId & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSetnew.Fields("AED_TYPE").Value), "", objRecordSetnew.Fields("AED_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            'If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing 'AMIT
            If strTaxType = "AED" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("AED_AMOUNT").Value), 0, objRecordSetnew.Fields("AED_AMOUNT").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("AED_PER").Value), 0, objRecordSetnew.Fields("AED_PER").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSetnew.Close()
                            objRecordSetnew = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '10792712 ENDED 

        'tcs changes
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("TCSTAXAMOUNT").Value), 0, objRecordSetnew.Fields("TCSTAXAMOUNT").Value)
        If blnTCSTax = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intTCSRoundOffDecimal)
        ElseIf blnTCSTax = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("TCSTax_Per").Value), 0, objRecordSetnew.Fields("TCSTax_Per").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            'TCS CHANGES
            strRetVal = GetTaxGlSl("TCS")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for TCS", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»TCS»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                If optDrnote.Checked = True Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»TCS for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & _
                                         dblTaxAmt & "»»TCS for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If

            End If
            iCtr = iCtr + 1
        End If

        'tcs changes

        ''Packing Posting
        Dim curPacking_per As Decimal
        curPacking_per = IIf(IsDBNull(objRecordSetnew.Fields("Packing").Value), 0, objRecordSetnew.Fields("Packing").Value)
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("Packing_Amount").Value), 0, objRecordSetnew.Fields("Packing_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If curPacking_per > 0 Then
            strRetVal = GetTaxGlSl("PKT")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for PACKING", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If dblBaseCurrencyAmount > 0 Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»PKT»0»" & "»»" & curPacking_per & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                iCtr = iCtr + 1
            End If
        End If
        'SST Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("Surcharge_Sales_Tax_Amount").Value), 0, objRecordSetnew.Fields("Surcharge_Sales_Tax_Amount").Value)
        If blnISSurChargeTaxRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intSSTRoundOffDecimal)
        ElseIf blnISSurChargeTaxRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("Surcharge_SalesTax_Per").Value), 0, objRecordSetnew.Fields("Surcharge_SalesTax_Per").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("SST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for SST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            iCtr = iCtr + 1
        End If
        '101188073
        'CGST
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("CGST_AMT").Value), 0, objRecordSetnew.Fields("CGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("CGST_PERCENT").Value), 0, objRecordSetnew.Fields("CGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("CGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for CGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»CGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                If optDrnote.Checked = True Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                     dblTaxAmt & "»»CGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & _
                                     dblTaxAmt & "»»CGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If

            End If

            iCtr = iCtr + 1
        End If
        'SGST
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("SGST_AMT").Value), 0, objRecordSetnew.Fields("SGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("SGST_PERCENT").Value), 0, objRecordSetnew.Fields("SGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("SGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for SGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                If optDrnote.Checked = True Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»sGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & _
                                         dblTaxAmt & "»»sGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If

            End If
            iCtr = iCtr + 1
        End If
        'UTGST
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("UTGST_AMT").Value), 0, objRecordSetnew.Fields("UTGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("UTGST_PERCENT").Value), 0, objRecordSetnew.Fields("UTGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("UTGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for UTGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»UTGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                If optDrnote.Checked = True Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                        dblTaxAmt & "»»UTGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & _
                                        dblTaxAmt & "»»UTGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
            End If



            iCtr = iCtr + 1
        End If
        'IGST
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("IGST_AMT").Value), 0, objRecordSetnew.Fields("IGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("IGST_PERCENT").Value), 0, objRecordSetnew.Fields("IGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("IGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for IGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»IGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                If optDrnote.Checked = True Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»IGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & _
                                         dblTaxAmt & "»»IGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If

            End If

            iCtr = iCtr + 1
        End If
        'CCESS
        dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("CCESS_AMT").Value), 0, objRecordSetnew.Fields("CCESS_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSetnew.Fields("COMPENSATION_CESS_PERCENT").Value), 0, objRecordSetnew.Fields("COMPENSATION_CESS_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("GSTEC")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for COMP.CESS", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSetnew.Close()
                    objRecordSetnew = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»GSTEC»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            iCtr = iCtr + 1
        End If
        '101188073
        'If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSetnew.Close() AMIT


        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSetnew.Close()
        objRecordSetnew = Nothing
        objRecordSetnew = New ADODB.Recordset
        objRecordSetnew.CursorLocation = ADODB.CursorLocationEnum.adUseClient


        If mblnSUPP_INVOICE_ONEITEM_MULTIPLE = True Then
            objRecordSetnew.Open("SELECT sum(SupplementaryInv_dtl.basic_amountdiff )as basic_amount , item_mst.GlGrp_code ,SupplementaryInv_dtl.totalexciseamount as totalexciseamount FROM SupplementaryInv_Hdr, item_mst, SupplementaryInv_dtl  WHERE SupplementaryInv_Hdr.Unit_code=item_mst.Unit_code and SupplementaryInv_Hdr.unit_code=SupplementaryInv_dtl.unit_code and SupplementaryInv_Hdr.doc_no =SupplementaryInv_dtl.doc_no  and SupplementaryInv_Hdr.Doc_No='" & docno & "' and SupplementaryInv_dtl.Item_Code=item_mst.Item_Code and SupplementaryInv_Hdr.Location_Code='" & gstrUNITID & "' and SupplementaryInv_Hdr.Unit_Code='" & gstrUNITID & "' and supplementaryinv_hdr.doc_no='" & docno & "' group by item_mst.GlGrp_code,SupplementaryInv_dtl.totalexciseamount ", mP_Connection)
        Else
            objRecordSetnew.Open("SELECT sum(SupplementaryInv_dtl.basic_amountdiff )as basic_amount , item_mst.GlGrp_code ,SupplementaryInv_dtl.totalexciseamount as totalexciseamount FROM SupplementaryInv_Hdr, item_mst, SupplementaryInv_dtl  WHERE SupplementaryInv_Hdr.Unit_code=item_mst.Unit_code and SupplementaryInv_Hdr.unit_code=SupplementaryInv_dtl.unit_code and SupplementaryInv_Hdr.doc_no =SupplementaryInv_dtl.doc_no  and SupplementaryInv_Hdr.Doc_No='" & docno & "' and SupplementaryInv_Hdr.Item_Code=item_mst.Item_Code and SupplementaryInv_Hdr.Location_Code='" & gstrUNITID & "' and SupplementaryInv_Hdr.Unit_Code='" & gstrUNITID & "' and supplementaryinv_hdr.doc_no='" & docno & "' group by item_mst.GlGrp_code,SupplementaryInv_dtl.totalexciseamount ", mP_Connection)
        End If

        If objRecordSetnew.EOF Then
            MsgBox("Item details not found.", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSetnew.Close()
                objRecordSetnew = Nothing
            End If
            Exit Function
        End If
        While Not objRecordSetnew.EOF
            strGlGroupId = Trim(IIf(IsDBNull(objRecordSetnew.Fields("GlGrp_code").Value), "", objRecordSetnew.Fields("GlGrp_code").Value))
            'Basic Amount Posting
            dblBasicAmount = IIf(IsDBNull(objRecordSetnew.Fields("Basic_Amount").Value), 0, objRecordSetnew.Fields("Basic_Amount").Value)
            If blnISBasicRoundOff = False Then
                dblBasicAmount = System.Math.Round(dblBasicAmount, intBasicRoundOffDecimal)
            ElseIf blnISBasicRoundOff = True Then
                dblBasicAmount = System.Math.Round(dblBasicAmount, 0)
            End If
            If dblBasicAmount > 0 Then
                'initializing the item gl and sl************************
                strRetVal = GetItemGLSL(strGlGroupId, mstrPurposeCode)
                If strRetVal = "N" Then
                    CreateStringForAccounts = False
                    If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSetnew.Close()
                        objRecordSetnew = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetVal, "»")
                strItemGL = varTmp(0)
                strItemSL = varTmp(1)
                'initializing of item gl and sl ends here****************
                'rsSalesInvType.ResultSetClose()
                rsSalesInvType = New ClsResultSetDB
                rsSalesInvType.GetResult("Select Invoice_type,Sub_category from SalesChallan_dtl where Unit_code='" & gstrUNITID & "' and Doc_no ='" & docno & "'")
                strInvoiceType = rsSalesInvType.GetValue("Invoice_type")
                strInvoiceSubType = rsSalesInvType.GetValue("Sub_category")
                rsSalesInvType.ResultSetClose()
                rsSalesInvType = Nothing
                'Posting the basic amount into cost centers, percentage wise

                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                objTmpRecordset = Nothing
                objTmpRecordset = New ADODB.Recordset
                objTmpRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient

                objTmpRecordset.Open("SELECT Location_Code,Invoice_Type,Sub_Type,ccM_ccCode,ccM_cc_percentage FROM invcc_dtl WHERE Unit_code='" & gstrUNITID & "' and Invoice_Type='" & strInvoiceType & "' AND Sub_Type = '" & strInvoiceSubType & "' AND Location_Code ='" & gstrUNITID & "' AND ccM_cc_Percentage > 0", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    While Not objTmpRecordset.EOF
                        dblCCShare = (dblBasicAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value

                        If gblnGSTUnit = False Then
                            'mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSetnew.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & "ABC" & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                        Else
                            If optDrnote.Checked = True Then
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & dblBasicAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»DR»" & dblBasicAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If

                        End If

                        objTmpRecordset.MoveNext()
                        iCtr = iCtr + 1
                    End While
                Else
                    If gblnGSTUnit = False Then
                        'mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSetnew.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        If gstrUNITID = "STH" Then
                            If optDrnote.Checked = True Then
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & dblBasicAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»DR»" & dblBasicAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If

                        Else
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & "XYZ" & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        End If

                    Else
                        If optDrnote.Checked = True Then
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & dblBasicAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»DR»" & dblBasicAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If

                    End If

                    iCtr = iCtr + 1
                End If
            End If
            '*********************************************************
            '*********************************************************
            ''EXC Duty Posting
            '*********************************************************
            '*********************************************************
            dblTaxAmt = IIf(IsDBNull(objRecordSetnew.Fields("TotalExciseAmount").Value), 0, objRecordSetnew.Fields("TotalExciseAmount").Value)
            If blnISExciseRoundOff = False Then
                dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intExciseRoundOffDecimal)
            ElseIf blnISExciseRoundOff = True Then
                dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
            End If
            If dblBaseCurrencyAmount > 0 Then
                'initializing the tax gl and sl here
                strRetVal = GetTaxGlSl("EXC")
                If strRetVal = "N" Then
                    MsgBox("GL for ARTAX is not defined for EXC", MsgBoxStyle.Information, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSetnew.Close()
                        objRecordSetnew = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetVal, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»EXC»0»" & Trim(objRecordSetnew.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                iCtr = iCtr + 1
            End If
            objRecordSetnew.MoveNext()
        End While
        ''Posting of rounded off amount

        strRetVal = GetItemGLSL("", "Rounded_Amt")
        If strRetVal = "N" Then
            CreateStringForAccounts = False
            If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSetnew.Close()
                objRecordSetnew = Nothing
            End If
            Exit Function
        End If
        varTmp = Split(strRetVal, "»")
        strItemGL = varTmp(0)
        strItemSL = varTmp(1)
        dblBaseCurrencyAmount = dblInvoiceAmtRoundOff_diff
        dblBaseCurrencyAmount = System.Math.Round(dblBaseCurrencyAmount, intBasicRoundOffDecimal)
        If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»RND»0»" & "»»0»" & strItemGL & "»" & strItemSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»"
            If dblBaseCurrencyAmount < 0 Then
                mstrDetailString = mstrDetailString & "Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "Dr»»»»»»0»0»0»0»0" & "¦"
            End If
        End If
        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSetnew.Close()
            objRecordSetnew = Nothing
        End If
        CreateStringForAccounts = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        CreateStringForAccounts = False
        If objRecordSetnew.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSetnew.Close()
            objRecordSetnew = Nothing
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetItemGLSL(ByVal InventoryGlGroup As String, ByVal PurposeCode As String) As String
        Dim objRecordSet As New ADODB.Recordset
        Dim strGL As String
        Dim strSL As String
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE Unit_code='" & gstrUnitId & "' and invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE Unit_code='" & gstrUnitId & "' and gbl_prpsCode = '" & PurposeCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRecordSet.EOF Then
                GetItemGLSL = "N"
                MsgBox("GL and SL not defined for Purpose Code: " & PurposeCode, MsgBoxStyle.Information, ResolveResString(100))
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            Else
                strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_glCode").Value), "", objRecordSet.Fields("gbl_glCode").Value))
                strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_slCode").Value), "", objRecordSet.Fields("gbl_slCode").Value))
            End If
        Else
            strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_glcode").Value), "", objRecordSet.Fields("invGld_glcode").Value))
            strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_slcode").Value), "", objRecordSet.Fields("invGld_slcode").Value))
        End If
        If strGL = "" Then
            GetItemGLSL = "N"
            MsgBox("GL and SL not defined for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, ResolveResString(100))
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetItemGLSL = strGL & "»" & strSL
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetItemGLSL = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where Unit_code='" & gstrUnitId & "' and tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            GetTaxGlSl = "N"
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetTaxGlSl = Trim(IIf(IsDBNull(objRecordSet.Fields("tx_glCode").Value), "", objRecordSet.Fields("tx_glCode").Value)) & "»" & Trim(IIf(IsDBNull(objRecordSet.Fields("tx_slCode").Value), "", objRecordSet.Fields("tx_slCode").Value))
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetTaxGlSl = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        Dim strdeletehdrsql As String = String.Empty
        Dim strdeletedtlsql As String = String.Empty
        Dim strsuppdetailsql As String = String.Empty
        Dim strsuppdetailsqlbatchwise As String = String.Empty
        Dim strinvoiceno As String
        Dim intmainloop As Integer
        Dim blnnoitemselected As Boolean = False
        Try
            If OptReprint.Checked = True Then
                MsgBox("Can't Delete Locked invoice!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Exit Sub
            Else
                With fsprDtls
                    strdeletehdrsql = ""
                    strdeletedtlsql = ""
                    strsuppdetailsql = ""
                    strsuppdetailsqlbatchwise = ""
                    If .MaxRows > 0 Then
                        For intmainloop = 1 To .MaxRows

                            .Row = intmainloop
                            .Col = ENUMGRIDDTLS.ENUM_Mark
                            If CBool(.Value) = True Then
                                blnnoitemselected = True
                                strinvoiceno = ""
                                .Row = intmainloop
                                .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO
                                strinvoiceno = .Value
                                strdeletehdrsql += "delete from supplementaryinv_hdr where unit_code='" & gstrUnitId & "' and bill_flag=0 and doc_no ='" & strinvoiceno & "'"
                                strdeletedtlsql += "delete from supplementaryinv_dtl where unit_code='" & gstrUNITID & "' and doc_no ='" & strinvoiceno & "'"
                                strsuppdetailsql += "delete from supplementaryData_detail where unit_code='" & gstrUNITID & "' and invoice_no ='" & strinvoiceno & "'"
                                strsuppdetailsqlbatchwise += "delete from supplementaryData_batchwise where unit_code='" & gstrUNITID & "' and invoice_no ='" & strinvoiceno & "'"
                            End If
                        Next
                        If blnnoitemselected = False Then
                            MsgBox("No Invoice has been selected ", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If
                        If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                            mP_Connection.BeginTrans()
                            If Len(Trim(strsuppdetailsql)) > 0 Then
                                mP_Connection.Execute(strsuppdetailsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If Len(Trim(strsuppdetailsqlbatchwise)) > 0 Then
                                mP_Connection.Execute(strsuppdetailsqlbatchwise, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If Len(Trim(strdeletehdrsql)) > 0 And Len(Trim(strdeletedtlsql)) > 0 Then
                                mP_Connection.Execute(strdeletehdrsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute(strdeletedtlsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.CommitTrans()
                                MsgBox("Invoice deleted Successfully", MsgBoxStyle.Information, ResolveResString(100))
                                Call BtnFetchshow_Click(Btnshow, New System.EventArgs())
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If

                    Else
                        MsgBox("No Row has been selected ", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If

                End With
            End If

        Catch ex As Exception
            mP_Connection.RollbackTrans()
            RaiseException(ex)
        End Try
    End Sub
    Private Sub optCheckAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCheckAll.CheckedChanged
        If eventSender.Checked Then
            With fsprDtls
                If .MaxRows > 0 Then
                    For lngCounter = 1 To .MaxRows
                        .Row = lngCounter : .Col = ENUMGRIDDTLS.ENUM_Mark : .Value = CStr(1)
                    Next
                End If
            End With
        End If
    End Sub
    Private Sub optUncheckAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optunCheckAll.CheckedChanged
        If eventSender.Checked Then
            With fsprDtls
                If .MaxRows > 0 Then
                    For lngCounter = 1 To .MaxRows
                        .Row = lngCounter : .Col = ENUMGRIDDTLS.ENUM_Mark : .Value = CStr(0)
                    Next
                End If
            End With
        End If
    End Sub

    Private Sub btn_Lock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Lock.Click
        Dim rsInvoiceType As ClsResultSetDB
        Dim strRetVal As String
        Dim objDrCr As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim strRefInvoiceNo As String
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intNoCopies As Short
        Dim strAccountCode As String = String.Empty
        Dim strVoucher As String = String.Empty
        Dim rstemp As ClsResultSetDB
        Dim intloop As Integer
        Dim objTmpRecordset1 As New ADODB.Recordset
        Dim currentinvno As String
        Dim blnnoitemselected As Boolean = False
        Dim intmainloop As Integer
        Dim strinvoiceno As String
        Dim mstrfinalmsg As String
        Dim inttotalnoofinvoices As Integer
        Try
            If Me.OptReprint.Checked = True Then
                MsgBox("Select Generate Button ", MsgBoxStyle.Information, ResolveResString(100))
                fsprDtls.MaxRows = 0
                txtFromInvoice.Text = ""
                txtToInvoice.Text = ""
                Exit Sub

            End If
            With fsprDtls
                For intmainloop = 1 To .MaxRows
                    .Row = intmainloop
                    .Col = ENUMGRIDDTLS.ENUM_Mark

                    If CBool(.Value) = True Then
                        blnnoitemselected = True
                        Exit For
                    End If
                Next
            End With

            If fsprDtls.MaxRows > 0 Then
                mP_Connection.Execute("delete from TMP_SUPPLEMENTARY_BULKINVOICES where iP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If OptGenerate.Checked = True Then
                    If blnnoitemselected = True Then
                        If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                            mstrunlockedfinalmsg = ""
                            mstlockedfinalmsg = ""
                            With fsprDtls
                                For intmainloop = 1 To .MaxRows
                                    .Row = intmainloop
                                    .Col = ENUMGRIDDTLS.ENUM_Mark
                                    .Row = intmainloop : .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO : Ctlinvoice = Trim(.Text)
                                    '                                        If PostInvoice() = False Then
                                    mP_Connection.Execute("insert into TMP_SUPPLEMENTARY_BULKINVOICES (UNIT_CODE,DOC_NO,iP_ADDRESS) select '" & gstrUnitId & "','" & Ctlinvoice & "','" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    'End If
                                Next
                            End With
                            inttotalnoofinvoices = Find_Value("select count(distinct doc_no) from  TMP_SUPPLEMENTARY_BULKINVOICES  where unit_code='" & gstrUnitId & "' and ip_address='" & gstrIpaddressWinSck & "'")
                            With fsprDtls
                                If CBool(Find_Value("SELECT SEPARATE_SUPP_SERIES FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUnitId & "'")) = True Then
                                    mstrSuffix = Find_Value("SET DATEFORMAT 'DMY' SELECT ISNULL(SUFFIX_SUP,0) AS SUFFIX  from saleconf  WHERE UNIT_CODE='" + gstrUnitId + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='F' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE")
                                Else
                                    mstrSuffix = Find_Value("SET DATEFORMAT 'DMY' SELECT ISNULL(SUFFIX,0) AS SUFFIX  from saleconf  WHERE UNIT_CODE='" + gstrUnitId + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='F' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE")
                                End If
                                If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE SUPP_INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUnitId & "'") Then
                                    mblnsamelockingdate = True
                                End If
                                For intmainloop = 1 To .MaxRows
                                    .Row = intmainloop
                                    .Col = ENUMGRIDDTLS.ENUM_Mark
                                    If CBool(.Value) = True Then
                                        If intmainloop > 1 Then
                                            mstrstrprevinvoice = Ctlinvoice
                                        End If
                                        .Row = intmainloop : .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO : Ctlinvoice = Trim(.Text)
                                        If (mstrstrprevinvoice <> Ctlinvoice) Or intmainloop = 1 Then

                                            If PostInvoice(mstrSuffix) = False Then
                                            End If
                                        End If
                                    End If

                                Next
                            End With

                            If mstlockedfinalmsg.Length.ToString > 0 Then
                                mstlockedfinalmsg = mstlockedfinalmsg + " TO : " + mInvNo
                            Else
                            End If
                            mstrfinalmsg = mstlockedfinalmsg + vbCrLf + mstrunlockedfinalmsg
                            If mstrfinalmsg.Trim.ToString.Length > 0 Then
                                MsgBox(mstrfinalmsg, MsgBoxStyle.OkOnly, ResolveResString(100))
                                ''ADDED BY SUMIT KUMAR TO CREATE AND SAVE SUPPLEMENTARY PDF FILE ON 21 AUG 2019

                                Try
                                    If FLAGISCUSTOMER_PDF_REQ = False Then
                                        txtFromInvoice.Text = ""
                                        txtToInvoice.Text = ""
                                        fsprDtls.MaxRows = 0
                                        Exit Sub
                                    End If
                                    Dim invoice_series() As String
                                    invoice_series = mstrfinalmsg.Split(":")
                                    If invoice_series.Count > 2 And mstrunlockedfinalmsg.Trim() = "" Then
                                        If intmainloop > 1 Then
                                            txtFromInvoice.Text = ""
                                            txtToInvoice.Text = String.Empty
                                            txtFromInvoice.Text = Convert.ToString(Replace(invoice_series(1).ToString().Trim(), "TO", "")).Trim()
                                            txtToInvoice.Text = invoice_series(2).ToString().Trim()
                                            If (txtFromInvoice.Text.Trim = "" Or txtToInvoice.Text.Trim = "") Then
                                                Exit Sub
                                            End If
                                        Else
                                            txtFromInvoice.Text = mInvNo.ToString()
                                            txtFromInvoice.Text = mInvNo.ToString()
                                            If (txtFromInvoice.Text.Trim = "" Or txtToInvoice.Text.Trim = "") Then
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                    lblpdfcreate.Visible = True
                                    btnSavePDFBinary_Click(btnSavePDFBinary, New System.EventArgs())
                                Catch ex As Exception
                                    lblpdfcreate.Visible = False
                                End Try

                                ''ENDED 
                                lblpdfcreate.Visible = False
                                txtFromInvoice.Text = ""
                                txtToInvoice.Text = ""
                                fsprDtls.MaxRows = 0
                                Exit Sub
                            End If

                        End If
                    Else
                        MsgBox("No Row has been selected ", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                End If
            End If
            Call BtnFetchshow_Click(Btnshow, New System.EventArgs())
            lblpdfcreate.Visible = False
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub



    Private Function PostInvoice(ByVal pstrSuffix As String) As Boolean
        Dim rsInvoiceType As ClsResultSetDB
        Dim strRetVal As String
        Dim objDrCr As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim strRefInvoiceNo As String
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intNoCopies As Short
        Dim strAccountCode As String = String.Empty
        Dim strVoucher As String = String.Empty
        Dim rstemp As ClsResultSetDB
        Dim intloop As Integer
        Dim objTmpRecordset1 As New ADODB.Recordset
        Dim currentinvno As String
        Dim blnnoitemselected As Boolean = False
        Dim intmainloop As Integer
        Dim strinvoiceno As String
        Dim strfirstinvoice As String
        Dim strlastinvoice As String
        Dim strUpdate As New StringBuilder
        Dim strfinarinvmapping As String


        Try

            'If IsNothing(MySQLConn) = True Then
            'MySQLConn = New SqlConnection()
            'MySQLConn = SqlConnectionclass.GetConnection()
            'End If

            SqlConnectionclass.BeginTrans()
            mInvNo = CInt(SqlConnectionclass.ExecuteScalar("SET DATEFORMAT 'DMY' select isnull(CURRENT_NO_SUPPINVOICE,0)+1   from saleconf  WHERE UNIT_CODE='" + gstrUnitId + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='F' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"))
            'If CBool(Find_Value("SELECT SEPARATE_SUPP_SERIES FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'")) = True Then
            'mstrSuffix = Find_Value("SET DATEFORMAT 'DMY' SELECT ISNULL(SUFFIX_SUP,0) AS SUFFIX  from saleconf  WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='F' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE")
            'Else
            'mstrSuffix = Find_Value("SET DATEFORMAT 'DMY' SELECT ISNULL(SUFFIX,0) AS SUFFIX  from saleconf  WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='F' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE")
            'End If

            mblnSameSeries = "1"
            mstrPurposeCode = Convert.ToString(SqlConnectionclass.ExecuteScalar("SET DATEFORMAT 'DMY' select inv_GLD_prpsCode from saleconf  WHERE UNIT_CODE='" + gstrUnitId + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='F' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"))
            mSaleConfNo = mInvNo
            mInvNo = CStr(pstrSuffix) + CStr(mInvNo)

            If mblnpostinfin = True Then
                If Not CreateStringForAccounts(Ctlinvoice) Then
                    Exit Function
                End If
            End If
            strUpdate.AppendLine("update SupplementaryInv_dtl set Doc_no = '" & mInvNo & "' where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Ctlinvoice & "'")
            If mblnsamelockingdate = True Then
                strUpdate.AppendLine("update SupplementaryInv_hdr set invoice_date=Convert(varchar(12), getdate(), 106) ,Doc_no = '" & mInvNo & "', Bill_flag = 1 where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Ctlinvoice & "'")
            Else
                strUpdate.AppendLine("update SupplementaryInv_hdr set Doc_no = '" & mInvNo & "', Bill_flag = 1 where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Ctlinvoice & "'")
            End If
            strUpdate.AppendLine("Update dbo.SuppCreditAdvise_Dtl set Doc_no = '" & mInvNo & "' where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Ctlinvoice & "'")
            If mblnsamelockingdate = True Then
                strUpdate.AppendLine("Update MarutisupplementaryData set invoice_date=Convert(varchar(12), getdate(), 106),invoice_no= '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")
                strUpdate.AppendLine("Update supplementaryData_detail set invoice_date=Convert(varchar(12), getdate(), 106),invoice_no= '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")

                strUpdate.AppendLine("Update MarutisupplementaryData_batchwise set invoice_date=Convert(varchar(12), getdate(), 106),invoice_no= '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")
                strUpdate.AppendLine("Update supplementaryData_batchwise set invoice_date=Convert(varchar(12), getdate(), 106),invoice_no= '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")
            Else
                strUpdate.AppendLine("Update MarutisupplementaryData set invoice_no = '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")
                strUpdate.AppendLine("Update supplementaryData_detail set invoice_date=Convert(varchar(12), getdate(), 106),invoice_no= '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")
                strUpdate.AppendLine("Update MarutisupplementaryData_batchwise set invoice_no = '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")
                strUpdate.AppendLine("Update supplementaryData_batchwise set invoice_date=Convert(varchar(12), getdate(), 106),invoice_no= '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and invoice_no = '" & Ctlinvoice & "'")
            End If


            If Not mblnSameSeries Then
                strUpdate.AppendLine("update saleconf set CURRENT_NO_SUPPINVOICE = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Invoice_type = '" & strInvoiceType & "' and Location_Code='" & gstrUnitId & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE")
            Else
                strUpdate.AppendLine("update saleconf set CURRENT_NO_SUPPINVOICE = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Single_Series = 1 and Location_Code='" & gstrUnitId & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE")
            End If
            'added by priti on 29.04.2019 to solve problem of data not inserting into finance table
            If strUpdate.ToString.Trim.Length > 0 Then
                SqlConnectionclass.ExecuteNonQuery(strUpdate.ToString())
            End If
            ' priti code ends here

            If mblnpostinfin = True Then
                If gstrUNITID <> "STH" Then
                    If optDrnote.Checked = True Then
                        prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DRG"
                    Else
                        prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "CRG"
                    End If
                End If
                strRetVal = objDrCr.SetARDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                    prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""

                    If Mid(strRetVal, 1, 1) = "Y" Then
                        strVoucher = Mid(strRetVal, 16, 12)

                        If optDrnote.Checked = True Then
                            strfinarinvmapping = "INSERT INTO FIN_AR_INV_MAPPING(VO_NO,UNIT_CODE,INV_NO,SO_NO,ITEM_CODE,HSN_CODE,QTY," &
                    "Rate,CGST_PER,CGST_AMT,SGST_PER,SGST_AMT,IGST_PER,IGST_AMT,BASIC_AMT,TAX_AMT,TOTAL_AMT,STATUS, " &
                    "DRCRNOTE, AR_AP,DR_CR,TCS_RT,TCS_PER,TCS_AMT,PREVRATE,SERIALNO ) SELECT '" & strVoucher & "','" & gstrUNITID & "','" & mInvNo & "',h.cust_ref,d.item_code,h.hsn_sac_code,0,d.rate_diff," &
                    "d.CGST_PERCENT,d.DIFF_CGST_AMT,d.sGST_PERCENT,d.DIFF_sGST_AMT,d.IGST_PERCENT,d.DIFF_IGST_AMT ,D.BASIC_AMOUNTDIFF,D.BASIC_AMOUNTDIFF,d.total_amountdiff, " &
                    "'A','" & strVoucher & "','AR','DR' ,d.TCSTax_Type,isnull(d.TCSTax_Per,0),d.TCSTaxAmount ,d.PREVRATE,isnull(d.serialno,'')   from supplementaryinv_hdr h (nolock) ,supplementaryinv_dtl d (nolock) " &
                    "where h.unit_code=d.unit_code and h.doc_no=d.doc_no AND d.unit_code='" & gstrUNITID & "' and d.doc_no='" & mInvNo & "' and  " &
                    " H.INVOICE_DATE >=( SELECT TOP 1 FIN_START_DATE FROM SALECONF WHERE UNIT_CODE='" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE )AND  " &
                    " D.SUPPINVDATE>=( SELECT TOP 1 FIN_START_DATE FROM SALECONF WHERE UNIT_CODE='" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE )AND  " &
                    "NOT EXISTS ( SELECT TOP 1 VO_NO FROM FIN_AR_INV_MAPPING FINMAP (NOLOCK) " &
                    "WHERE FINMAP.VO_NO='" & strVoucher & "' " &
                    "AND FINMAP.UNIT_CODE='" & gstrUNITID & "' " &
                    "AND FINMAP.dr_cr ='DR' " &
                    "AND FINMAP.AR_AP ='AR') "

                        Else
                            strfinarinvmapping = "INSERT INTO FIN_AR_INV_MAPPING(VO_NO,UNIT_CODE,INV_NO,SO_NO,ITEM_CODE,HSN_CODE,QTY," &
                    "Rate,CGST_PER,CGST_AMT,SGST_PER,SGST_AMT,IGST_PER,IGST_AMT,BASIC_AMT,TAX_AMT,TOTAL_AMT,STATUS, " &
                    "DRCRNOTE, AR_AP,DR_CR,TCS_RT,TCS_PER,TCS_AMT,PREVRATE,SERIALNO) SELECT '" & strVoucher & "','" & gstrUNITID & "','" & mInvNo & "',h.cust_ref,d.item_code,h.hsn_sac_code,0,d.rate_diff," &
                    "d.CGST_PERCENT,d.DIFF_CGST_AMT,d.sGST_PERCENT,d.DIFF_sGST_AMT,d.IGST_PERCENT,d.DIFF_IGST_AMT ,D.BASIC_AMOUNTDIFF,D.BASIC_AMOUNTDIFF,d.total_amountdiff, " &
                    "'A','" & strVoucher & "','AR','CR',d.TCSTax_Type,isnull(d.TCSTax_Per,0),d.TCSTaxAmount ,d.PREVRATE,isnull(d.serialno,'') from supplementaryinv_hdr h (nolock) ,supplementaryinv_dtl d (nolock) " &
                    "where h.unit_code=d.unit_code and h.doc_no=d.doc_no and d.unit_code='" & gstrUNITID & "'  AND d.doc_no='" & mInvNo & "' and  " &
                    " H.INVOICE_DATE >=( SELECT TOP 1 FIN_START_DATE FROM SALECONF WHERE UNIT_CODE='" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE ) AND " &
                    " D.SUPPINVDATE>=( SELECT TOP 1 FIN_START_DATE FROM SALECONF WHERE UNIT_CODE='" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE )AND  " &
                    " NOT EXISTS ( SELECT TOP 1 VO_NO FROM FIN_AR_INV_MAPPING FINMAP (NOLOCK) " &
                    "WHERE FINMAP.VO_NO='" & strVoucher & "' " &
                    "AND FINMAP.UNIT_CODE='" & gstrUNITID & "' " &
                    "AND FINMAP.dr_cr ='CR' " &
                    "AND FINMAP.AR_AP ='AR') "
                        End If
                        If strfinarinvmapping.ToString.Trim.Length > 0 Then
                            SqlConnectionclass.ExecuteNonQuery(strfinarinvmapping)
                        End If

                    End If

                    strRetVal = CheckString(strRetVal)
                Else
                    strRetVal = "Y"
            End If
            'added by priti on 29.04.2019 to solve problem of data not inserting into finance table
            'Dim strQuery = "update supplementaryinv_hdr set VOUCHER_NO = '" & strVoucher & "' where Unit_code='" & gstrUnitId & "' and doc_no = '" & mInvNo & "'"
            Dim strQuery As String
            strQuery = "update supplementaryinv_hdr set VOUCHER_NO = '" & strVoucher & "' where Unit_code='" & gstrUNITID & "' and doc_no = '" & mInvNo & "'"
            strQuery += " and invoice_date >=( SELECT TOP 1 FIN_START_DATE FROM SALECONF WHERE UNIT_CODE='" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE ) "

            If strQuery.ToString.Trim.Length > 0 Then
                SqlConnectionclass.ExecuteNonQuery(strQuery)
            End If
            'priti code ends here


            If strQuery.ToString.Trim.Length > 0 Then
                SqlConnectionclass.ExecuteNonQuery(strQuery.ToString())
            End If

            If Not strRetVal = "Y" Then
                mstrunlockedfinalmsg = mstrunlockedfinalmsg + "Invoice no : " + Ctlinvoice + " Reason : " + strRetVal + vbCrLf
                SqlConnectionclass.RollbackTran()
            Else
                If mstlockedfinalmsg.ToString.Length = 0 Then
                    mstlockedfinalmsg = mstlockedfinalmsg + "Supplementary invoice Locked :  " + mInvNo
                Else
                    strlastinvoice = mInvNo
                End If
                SqlConnectionclass.CommitTran()
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
    Private Function DataExecution(ByVal strcustomercode As Object, ByVal strvoucher As Object) As Boolean
        Dim oCmd As New SqlCommand()

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

            DataExecution = True

            With oCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_FIN_GET_GST_AR_NOTE_DTL"

                .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 20).Value = strvoucher
                .Parameters.Add("@AR_AP", SqlDbType.VarChar, 6).Value = "AR"
                .Parameters.Add("@IP_ADD", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                .ExecuteNonQuery()
            End With

        Catch ex As Exception
            DataExecution = False
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function

    Public Function Find_Value(ByRef strField As String) As String
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then
            If IsDBNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub btn_print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_print.Click
        Dim minvno As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intNoCopies As Short
        Dim intmainloop As Short
        Dim strinvoiceno As String = ""
        Dim strprevinvoiceno As String = ""
        Dim BLNselectanyinvoice As String = False
        Dim strsql As String
        Dim strvouchernumber As Object
        Dim strRepPath As String
        Dim COPYNAME As String
        Dim strPath As String = String.Empty
        Dim strinvPath As String = String.Empty
        Dim strInvoiceDate As String = String.Empty
        Dim strCustVendorCode As String = String.Empty
        Dim strSuppInvNameAsPerCustomer As String = String.Empty

        '-----------------KILL PDF PRINTING PROCESS
        Try
            Dim aProcess As System.Diagnostics.Process
            aProcess = System.Diagnostics.Process.GetProcessById(pdfPrintProcID)
            If aProcess.HasExited = False Then
                aProcess.Kill()
            End If
        Catch ex As Exception
        End Try
        '-----------------KILL PDF PRINTING PROCESS

        Try
            If OptReprint.Checked = True Then

                frmRpt = New eMProCrystalReportViewer
                CR = New ReportDocument
                CR = frmRpt.GetReportDocument

                With fsprDtls
                    If .MaxRows <= 0 Then
                        MsgBox("No Row has been selected ", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    For intmainloop = 1 To .MaxRows

                        .Row = intmainloop
                        .Col = ENUMGRIDDTLS.ENUM_Mark
                        If CBool(.Value) = True Then
                            strinvoiceno = ""
                            .Row = intmainloop
                            .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO
                            strinvoiceno = .Text
                            .Col = ENUMGRIDDTLS.ENUM_AR_DEBITNOTE
                            strvouchernumber = Nothing
                            strvouchernumber = .Text
                            strcustomercode = Nothing
                            .Col = ENUMGRIDDTLS.ENUM_CUSTOMERCODE
                            strcustomercode = .Text

                            strPath = Find_Value("select Invoice_PDFCOPYPATH from CUSTOMER_MST  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strcustomercode & "'")
                            If OptOriginalcopy.Checked = True Then
                                mintnocopies = CInt(Find_Value("select ISNULL(NOOFCOPIES_SUPPINVOICE,0) NOOFCOPIES_SUPPINVOICE from CUSTOMER_MST  WHERE UNIT_CODE='" + gstrUnitId + "'AND CUSTOMER_CODE='" & strcustomercode & "'"))
                            ElseIf optCustomerCopy.Checked = True Then
                                mintnocopies = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUnitId + "'AND CUSTOMER_CODE='" & strcustomercode & "' AND ORIGINAL_REPRINT='O'"))
                            Else
                                If strPath.ToString.Length <= 0 Then
                                    MsgBox("Please Define Path in customer Master !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                    Exit Sub
                                End If
                                If Directory.Exists(strPath) = False Then
                                    Directory.CreateDirectory(strPath)
                                End If
                                strinvPath = strPath + "\" + strinvoiceno + ".PDF"
                                If strprevinvoiceno <> strinvoiceno Then
                                    If File.Exists(strinvPath) = True Then
                                        Kill(strinvPath)
                                    End If
                                End If
                                mintnocopies = CInt(Find_Value("select ISNULL(NOOFCOPIES_SUPPINVOICE,0) NOOFCOPIES_SUPPINVOICE from CUSTOMER_MST  WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strcustomercode & "'"))
                            End If

                            '' Start Added By Praveen For Name Change of SUPP Invoice in the Formate: DM089W_D012400001_22012024
                            strInvoiceDate = Find_Value("SELECT REPLACE(CONVERT(CHAR(10), INVOICE_DATE, 103), '/', '') [INVOICE_DATE] FROM SUPPLEMENTARYINV_HDR WHERE VOUCHER_NO='" + strvouchernumber + "' AND UNIT_CODE='" + gstrUNITID + "'")
                            strCustVendorCode = Find_Value("SELECT CUST_VENDOR_CODE FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" + strcustomercode + "' AND UNIT_CODE='" + gstrUNITID + "'")
                            strSuppInvNameAsPerCustomer = Find_Value("SELECT ISNULL(isSuppInvNameAsPerCustomer,0) FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" + strcustomercode + "' AND UNIT_CODE='" + gstrUNITID + "'")
                            If Convert.ToBoolean(strSuppInvNameAsPerCustomer) = True Then
                                strinvPath = strPath + "\" + strCustVendorCode + "_" + Convert.ToString(strvouchernumber) + "_" + strInvoiceDate + ".PDF"
                            End If
                            '' END

                            If mblnEwayFunctionality Then
                                Call IRN_QRBarcode(strvouchernumber)
                            End If

                            If optDrnote.Checked = True Then
                                If intmainloop = 1 Or strprevinvoiceno <> strinvoiceno Then
                                    strprevinvoiceno = strinvoiceno
                                    BLNselectanyinvoice = True
                                    If InvoiceGeneration(strinvoiceno) = True Then
                                        intMaxLoop = mintnocopies
                                        For intLoopCounter = 1 To intMaxLoop
                                            If OptReprint.Checked = True Then
                                                COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUnitId + "' AND CUSTOMER_CODE='" & strcustomercode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                            Else
                                                COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUnitId + "' AND CUSTOMER_CODE='" & strcustomercode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                            End If
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'" & COPYNAME & "'"
                                            frmRpt.SetReportDocument()

                                            If OptPDFCopy.Checked = True Then

                                                If mblnISTrueSignRequired Then
                                                    Dim strRESULT As Collections.ArrayList

                                                    CR.Export(GetExportOptions(strinvoiceno, strinvPath))
                                                    Dim OBJPdfConfig As Object = SqlConnectionclass.ExecuteScalar("SELECT COUNT(*) FROM INVOICE_PDF_CONFIG (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" + strcustomercode + "' AND INVOICE_TYPE='SUP' AND INVOICE_SUB_TYPE='S' And IS_ACTIVE=1")
                                                    Dim OBJCommonDigital_EINVOICING_CONFIG As Object = SqlConnectionclass.ExecuteScalar("select Count(*) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" + gstrUNITID + "' and CUSTOMER_CODE='" + strcustomercode + "' AND INVOICE_TYPE='SUP' and SUB_TYPE='S' AND IS_ACTIVE=1")
                                                    strRESULT = SavePDFInvoicesInDB(strcustomercode, strinvoiceno, "SUPPLEMENTARY", gstrUNITID, strinvPath, "SUP", "S", OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG, 0)
                                                    If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                                        File.WriteAllBytes(strinvPath, strRESULT.Item(1))
                                                    End If
                                                Else
                                                    CR.ExportToDisk(ExportFormatType.PortableDocFormat, strinvPath)
                                                End If
                                                
                                            Else
                                                If mblnISTrueSignRequired Then
                                                    Dim strRESULT As Collections.ArrayList

                                                    CR.Export(GetExportOptions(strinvoiceno, strinvPath))
                                                    Dim OBJPdfConfig As Object = SqlConnectionclass.ExecuteScalar("SELECT COUNT(*) FROM INVOICE_PDF_CONFIG (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" + strcustomercode + "' AND INVOICE_TYPE='SUP' AND INVOICE_SUB_TYPE='S' And IS_ACTIVE=1")
                                                    Dim OBJCommonDigital_EINVOICING_CONFIG As Object = SqlConnectionclass.ExecuteScalar("select Count(*) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" + gstrUNITID + "' and CUSTOMER_CODE='" + strcustomercode + "' AND INVOICE_TYPE='SUP' and SUB_TYPE='S' AND IS_ACTIVE=1")
                                                    strRESULT = SavePDFInvoicesInDB(strcustomercode, strinvoiceno, "SUPPLEMENTARY", gstrUNITID, strinvPath, "SUP", "S", OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG, 0)
                                                    If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                                        File.WriteAllBytes(strinvPath, strRESULT.Item(1))
                                                    End If

                                                    Dim MyProcess As New Process
                                                    Try
                                                        MyProcess.StartInfo.FileName = """" + mPdfReaderPath + """" ' + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                                                        MyProcess.StartInfo.Arguments = String.Format("{0}", " /t   " + """" + strinvPath + """")
                                                        MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                                        MyProcess.Start()
                                                        pdfPrintProcID = MyProcess.Id

                                                    Catch Ex As Exception
                                                        MsgBox(Ex.Message)
                                                    Finally
                                                    End Try

                                                Else
                                                    CR.PrintToPrinter(1, False, 0, 0)
                                                End If

                                            End If

                                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                                        Next
                                        'intMaxLoop = mintnocopies
                                        'For intLoopCounter = 1 To intMaxLoop
                                        '    Select Case intLoopCounter
                                        '        Case 1
                                        '            If OptReprint.Checked Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                                        '            ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                        '            Else
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                        '            End If
                                        '        Case 2
                                        '            If OptReprint.Checked = True Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                                        '            ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                                        '            Else
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER '"

                                        '            End If
                                        '        Case 3
                                        '            If OptReprint.Checked = True Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"

                                        '            ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                        '            Else
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                        '            End If
                                        '        Case 4
                                        '            If OptReprint.Checked = True Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"

                                        '            ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"
                                        '            Else
                                        '                CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                        '            End If
                                        '    End Select
                                        '    frmRpt.SetReportDocument()
                                        '    CR.PrintToPrinter(1, False, 0, 0)
                                        '    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                                        'Next
                                    End If
                                End If
                            Else
                                If DataExecution(strcustomercode, strvouchernumber) = False Then
                                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                                    Exit Sub
                                End If

                                mstrReportFilename = "AR_Note_GST"


                                'strRepPath = My.Application.Info.DirectoryPath & strReportName
                                strRepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt"
                                'CR.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")

                                CR = New ReportDocument()
                                frmRpt = New eMProCrystalReportViewer()
                                CR = frmRpt.GetReportDocument()
                                CR.Load(strRepPath)

                                With CR
                                    .SetParameterValue("repmaintitle", "Receipt voucher (Illustrative)")
                                    .SetParameterValue("cname", gstrCOMPANY)
                                    .SetParameterValue("address1", Trim(gstr_WRK_ADDRESS1))
                                    .SetParameterValue("address2", Trim(gstr_WRK_ADDRESS2))

                                    .RecordSelectionFormula = "{TMP_FIN_GET_GST_AR_NOTE_DETAIL.IP_ADDRESS}='" & gstrIpaddressWinSck & "'  and {TMP_FIN_GET_GST_AR_NOTE_DETAIL.DOCM_UNIT}='" & gstrUnitId & "'"
                                End With
                                'frmRpt.Show()
                                frmRpt.SetReportDocument()
                                BLNselectanyinvoice = True
                                If intmainloop = 1 Or strprevinvoiceno <> strinvoiceno Then
                                    strprevinvoiceno = strinvoiceno
                                    If OptPDFCopy.Checked = True Then
                                        CR.ExportToDisk(ExportFormatType.PortableDocFormat, strinvPath)

                                    Else
                                        CR.PrintToPrinter(1, False, 0, 0)
                                    End If

                                    'CR.PrintToPrinter(1, False, 0, 0)
                                End If

                            End If
                        End If
                    Next
                    If BLNselectanyinvoice = False Then
                        MsgBox("Please select invoice's!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Exit Sub
                    End If
                    If OptPDFCopy.Checked = True Then
                        MsgBox("PDF File Saved : " & vbCrLf & "PATH ! " + strPath, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    End If

                End With
            Else
                MsgBox("Can't print unlocked Invoice !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Public Function InvoiceGeneration(ByVal invoiceno As String) As Boolean
        Dim rstemp As ClsResultSetDB
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim Commissionerate As String
        Dim strsql As String
        Dim strCompMst, DeliveredAdd As String
        Dim strSuffix As String
        Dim strRefInvoiceNo As String
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim strCustomerCode As String = String.Empty
        Dim strVoucher As String = String.Empty
        Dim TinNo As String
        Dim blnPrintTinNo As Boolean
        Dim SuppDeb_Print As Boolean
        strRefInvoiceNo = "0"
        Dim oCmd As ADODB.Command
        Dim strIPAddress As String
        Dim strsql_oneitemmultiple As String
        Dim BLNMORETHAN_ONEITEM_SUPPINV As Boolean
        Dim BLNINVOICE_CUSTREF As Boolean
        On Error GoTo Err_Handler
        rstemp = New ClsResultSetDB
        strIPAddress = gstrIpaddressWinSck

        If invoiceno = "" Then InvoiceGeneration = False : Exit Function
        rstemp.GetResult("Select top 1 1 from SupplementaryInv_hdr where Unit_code='" & gstrUnitId & "' and Doc_no = '" & invoiceno & "' and Location_code = '" & gstrUnitId & "' and supp_invdetail='" & "O" & "'")
        If rstemp.GetNoRows > 0 Then
            strInvoiceType = "Inv"
            strInvoiceSubType = "F"
        Else
            rstemp.ResultSetClose()
            rstemp = New ClsResultSetDB

            BLNMORETHAN_ONEITEM_SUPPINV = CBool(Find_Value("SELECT MORETHAN_ONEITEM_SUPPINV FROM customer_mst WHERE UNIT_CODE='" & gstrUNITID & "' and CUSTOMER_CODE IN (SELECT ACCOUNT_CODE from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_no = '" & invoiceno & "')"))
            BLNINVOICE_CUSTREF = CBool(Find_Value("SELECT SUPPINVOICE_CUSTREF FROM customer_mst WHERE UNIT_CODE='" & gstrUNITID & "' and CUSTOMER_CODE IN (SELECT ACCOUNT_CODE from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_no = '" & invoiceno & "')"))
            If UCase(Trim(GetPlantName)) = "HILEX" Then
                strsql_oneitemmultiple = "Select TOP 1 Ref_Invno as RefDoc_no from MarutisupplementaryData where Unit_code='" & gstrUNITID & "' and INVOICE_NO = '" & invoiceno & "' and unit_code= '" & gstrUNITID & "'"
                strsql_oneitemmultiple += " union Select TOP 1 Ref_Invno as RefDoc_no from  supplementaryData_detail where Unit_code='" & gstrUNITID & "' and INVOICE_NO = '" & invoiceno & "'"
                strsql_oneitemmultiple += " union Select top 1 RefDoc_No from SupplementaryInv_dtl where Unit_code='" & gstrUNITID & "' and Doc_no = '" & invoiceno & "' and Location_code = '" & gstrUNITID & "' and isnull(RefDoc_No,0)<>0 "

                rstemp.GetResult(strsql_oneitemmultiple)
            Else
                If (mblnSUPP_INVOICE_ONEITEM_MULTIPLE = True And BLNMORETHAN_ONEITEM_SUPPINV = True) Or (BLNINVOICE_CUSTREF = True) Then
                    strsql_oneitemmultiple = "Select TOP 1 Ref_Invno as RefDoc_no from MarutisupplementaryData where Unit_code='" & gstrUNITID & "' and INVOICE_NO = '" & invoiceno & "' and unit_code= '" & gstrUNITID & "'"
                    strsql_oneitemmultiple += " union Select TOP 1 Ref_Invno as RefDoc_no from  supplementaryData_detail where Unit_code='" & gstrUNITID & "' and INVOICE_NO = '" & invoiceno & "'"
                    strsql_oneitemmultiple += " union Select top 1 RefDoc_No from SupplementaryInv_dtl where Unit_code='" & gstrUNITID & "' and Doc_no = '" & invoiceno & "' and Location_code = '" & gstrUNITID & "' and isnull(RefDoc_No,0)<>0 "

                    rstemp.GetResult(strsql_oneitemmultiple)
                Else
                    rstemp.GetResult("Select RefDoc_no from SupplementaryInv_dtl where Unit_code='" & gstrUNITID & "' and Doc_no = '" & invoiceno & "' and Location_code = '" & gstrUNITID & "'")
                End If

            End If


            'If (mblnSUPP_INVOICE_ONEITEM_MULTIPLE = True And BLNMORETHAN_ONEITEM_SUPPINV = True) Or (BLNINVOICE_CUSTREF = True) Then
            '    strsql_oneitemmultiple = "Select TOP 1 Ref_Invno as RefDoc_no from MarutisupplementaryData where Unit_code='" & gstrUNITID & "' and INVOICE_NO = '" & invoiceno & "' and unit_code= '" & gstrUNITID & "'"
            '    strsql_oneitemmultiple += " union Select TOP 1 Ref_Invno as RefDoc_no from  supplementaryData_detail where Unit_code='" & gstrUNITID & "' and INVOICE_NO = '" & invoiceno & "'"

            '    rstemp.GetResult(strsql_oneitemmultiple)
            'Else
            '    rstemp.GetResult("Select RefDoc_no from SupplementaryInv_dtl where Unit_code='" & gstrUNITID & "' and Doc_no = '" & invoiceno & "' and Location_code = '" & gstrUNITID & "'")
            'End If

            If rstemp.GetNoRows > 0 Then
                rstemp.MoveFirst()
                strRefInvoiceNo = rstemp.GetValue("RefDoc_no")
                rstemp.ResultSetClose()
                rstemp = New ClsResultSetDB
                rstemp.GetResult("Select Invoice_type,Sub_category,Account_Code From SalesChallan_dtl where Unit_code='" & gstrUnitId & "' and Doc_no = '" & strRefInvoiceNo & "' and Location_code = '" & gstrUnitId & "'")
                strInvoiceType = rstemp.GetValue("Invoice_type")
                strInvoiceSubType = rstemp.GetValue("Sub_category")

                If Trim(strRefInvoiceNo + "") = String.Empty Then
                    MsgBox(" Base Invoice No is not available !! - Supp Invoice no :" + invoiceno, MsgBoxStyle.Information, ResolveResString(100))
                    InvoiceGeneration = False : Exit Function
                End If
                If UCase(strInvoiceType) = "UNKNOWN" Then
                    strInvoiceType = "Inv"
                    strInvoiceSubType = "F"
                End If
                strCustomerCode = rstemp.GetValue("Account_Code")
            Else
                MsgBox(" Invalid Invoice No.", MsgBoxStyle.Information, ResolveResString(100))
                InvoiceGeneration = False : Exit Function
            End If
        End If
        rstemp.ResultSetClose()
        If UCase(Trim(GetPlantName)) = "HILEX" Then
            If (gblnGSTUnit = True) And CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUnitId + "'")) = False Then
                oCmd = New ADODB.Command
                With oCmd
                    .ActiveConnection = mP_Connection
                    .CommandTimeout = 0
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "PRC_INVOICEPRINTING_HILEX"
                    .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUnitId))
                    .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, gstrUnitId))
                    .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, invoiceno))
                    .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, "SUP"))
                    .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, "S"))
                    .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                    .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                    MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    oCmd = Nothing
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                oCmd = Nothing
                strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUnitId & "'"
            End If
        Else
            If (strInvoiceType = "TRF") Then
                oCmd = New ADODB.Command
                With oCmd
                    .ActiveConnection = mP_Connection
                    .CommandTimeout = 0
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "PRC_INVOICEPRINTING_MATE"
                    .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUnitId))
                    .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, gstrUnitId))
                    .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, invoiceno))
                    .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, "SUP"))
                    .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, "S"))
                    .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                    .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                    MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    oCmd = Nothing
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                oCmd = Nothing
                strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUnitId & "'"
            Else
                'If ((UCase(Trim(GetPlantName)) = "MATM" Or UCase(Trim(GetPlantName)) = "MR1")) And CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUnitId + "'")) = False Then
                If (((gblnGSTUnit = True) And CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))) = False) Or (gstrUNITID = "STH") Then
                    oCmd = New ADODB.Command
                    With oCmd
                        .ActiveConnection = mP_Connection
                        .CommandTimeout = 0
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "PRC_INVOICEPRINTING_MATE"
                        .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, invoiceno))
                        .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, "SUP"))
                        .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, "S"))
                        .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                        .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With
                    If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                        MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                        oCmd = Nothing
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Function
                    End If
                    oCmd = Nothing
                    strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                End If
            End If
        End If

        rstemp = New ClsResultSetDB
        rstemp.GetResult("SELECT inv_GLD_prpsCode,Single_Series,SuppReport_filename,NoCopies  FROM SaleConf WHERE Unit_code='" & gstrUnitId & "' and Invoice_Type='" & strInvoiceType & "' AND Sub_Type ='" & strInvoiceSubType & "' AND Location_Code='" & gstrUnitId & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ")
        If rstemp.GetNoRows > 0 Then
            mstrPurposeCode = IIf(IsDBNull(rstemp.GetValue("inv_GLD_prpsCode")), "", Trim(rstemp.GetValue("inv_GLD_prpsCode")))
            mblnSameSeries = rstemp.GetValue("Single_Series")
            'mintnocopies = rstemp.GetValue("nocopies")
            mstrReportFilename = IIf(IsDBNull(rstemp.GetValue("SuppReport_filename")), "", Trim(rstemp.GetValue("SuppReport_filename")))
            If mstrPurposeCode = "" Then
                MsgBox("Please select a Purpose Code in Sales Configuration", MsgBoxStyle.Information, ResolveResString(100))
                mstrPurposeCode = ""
                InvoiceGeneration = False
                Exit Function
            End If
        Else
            MsgBox("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", MsgBoxStyle.Information, ResolveResString(100))
            InvoiceGeneration = False
            Exit Function
        End If
        rstemp.ResultSetClose()
        strCompMst = "Select Reg_NO,Ecc_No,Range_1,Phone,Fax,Email,PLA_No,LST_No,CST_No,Division,Commissionerate,Invoice_Rule,Tin_no from Company_Mst where Unit_code='" & gstrUnitId & "'"
        rstemp = New ClsResultSetDB
        rstemp.GetResult(strCompMst)
        If rstemp.GetNoRows = 1 Then
            RegNo = rstemp.GetValue("Reg_NO")
            EccNo = rstemp.GetValue("Ecc_No")
            Range = rstemp.GetValue("Range_1")
            Phone = rstemp.GetValue("Phone")
            Fax = rstemp.GetValue("Fax")
            EMail = rstemp.GetValue("Email")
            PLA = rstemp.GetValue("PLA_No")
            UPST = rstemp.GetValue("LST_No")
            CST = rstemp.GetValue("CST_No")
            Division = rstemp.GetValue("Division")
            Commissionerate = rstemp.GetValue("Commissionerate")
            Invoice_Rule = rstemp.GetValue("Invoice_Rule")
            TinNo = rstemp.GetValue("Tin_no")
        End If
        rstemp.ResultSetClose()

        If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUnitId + "'")) = False Then
            rstemp = New ClsResultSetDB
            rstemp.GetResult("Select ConsigneeDetails from Sales_parameter where Unit_code='" & gstrUnitId & "'")
            If rstemp.GetValue("ConsigneeDetails") = False Then
                rstemp.ResultSetClose()
                rstemp = New ClsResultSetDB
                rstemp.GetResult("Select a.* from Customer_Mst a, SupplementaryInv_hdr b where a.unit_code=b.unit_code and a.Customer_code = b.Account_code and b.Doc_No = " & invoiceno & " and b.Location_Code='" & gstrUnitId & "' and a.Unit_code='" & gstrUnitId & "'")
                If rstemp.GetNoRows > 0 Then
                    DeliveredAdd = Trim(rstemp.GetValue("Ship_address1"))
                    If Len(Trim(DeliveredAdd)) Then
                        DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rstemp.GetValue("Ship_address2"))
                    Else
                        DeliveredAdd = Trim(rstemp.GetValue("Ship_address2"))
                    End If
                End If
            Else
                rstemp.ResultSetClose()
                rstemp = New ClsResultSetDB
                rstemp.GetResult("Select ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3 from Saleschallan_dtl where Unit_code='" & gstrUnitId & "' and Doc_No = " & strRefInvoiceNo & " and Location_Code='" & gstrUnitId & "'")
                If rstemp.GetNoRows > 0 Then
                    DeliveredAdd = Trim(rstemp.GetValue("ConsigneeAddress1"))
                    If Len(Trim(DeliveredAdd)) Then
                        DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rstemp.GetValue("ConsigneeAddress2"))
                    Else
                        DeliveredAdd = Trim(rstemp.GetValue("ConsigneeAddress2"))
                    End If
                    If Len(Trim(DeliveredAdd)) Then
                        DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rstemp.GetValue("ConsigneeAddress3"))
                    Else
                        DeliveredAdd = Trim(rstemp.GetValue("ConsigneeAddress3"))
                    End If
                End If
            End If
            rstemp.ResultSetClose()
            Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2

            If (strInvoiceType = "TRF") Then
                strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUnitId & "'"
            Else
                'If ((UCase(Trim(GetPlantName)) = "MATM" Or UCase(Trim(GetPlantName)) = "MR1")) Then
                If gblnGSTUnit = True Or gstrUNITID = "STH" Then
                    strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                Else
                    strsql = "{SupplementaryInv_hdr.Unit_Code}='" & gstrUnitId & "' and {SupplementaryInv_hdr.Location_Code}='" & gstrUnitId & "' and {SupplementaryInv_Hdr.Doc_No} =" & invoiceno
                End If
            End If

            If mstrReportFilename = "" Then
                MsgBox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, ResolveResString(100))
                InvoiceGeneration = False
                Exit Function
            End If
            If (strInvoiceType = "TRF" Or strInvoiceType = "ITD") And IsGSTINSAME(strCustomerCode) = True Then
                CR.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_A4REPORTS.rpt")
            Else
                CR.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
            End If
            With frmRpt
                CR.DataDefinition.FormulaFields("Registration").Text = "'" & RegNo & "'"
                CR.DataDefinition.FormulaFields("ECC").Text = "'" & EccNo & "'"
                CR.DataDefinition.FormulaFields("Range").Text = "'" & Range & "'"
                CR.DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
                CR.DataDefinition.FormulaFields("CompanyAddress").Text = "'" & Address & "'"
                CR.DataDefinition.FormulaFields("Phone").Text = "'" & Phone & "'"
                CR.DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
                CR.DataDefinition.FormulaFields("EMail").Text = "'" & EMail & "'"
                CR.DataDefinition.FormulaFields("PLA").Text = "'" & PLA & "'"
                CR.DataDefinition.FormulaFields("UPST").Text = "'" & UPST & "'"
                CR.DataDefinition.FormulaFields("CST").Text = "'" & CST & "'"
                CR.DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
                CR.DataDefinition.FormulaFields("commissionerate").Text = "'" & Commissionerate & "'"
                CR.DataDefinition.FormulaFields("InvoiceRule").Text = "'" & Invoice_Rule & "'"
                CR.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
                '10726518
                'CR.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
                If mbln_SHIPPING_ADDRESS = True Then
                    CR.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
                Else
                    CR.DataDefinition.FormulaFields("Address2").Text = "''"
                End If
                '10726518

                CR.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & CStr(invoiceno) & "'"
                If OptReprint.Checked = True Or FLAGISCUSTOMER_PDF_REQ = True Then
                    strVoucher = Find_Value("SELECT TOP 1 LORRYNO FROM TMP_INVOICEPRINT WHERE UNIT_CODE='" & gstrUnitId & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'")
                End If
                CR.DataDefinition.FormulaFields("voucherno").Text = "'" & strVoucher & "'"
                CR.DataDefinition.FormulaFields("minvno").Text = "'" & CStr(invoiceno) & "'"


                blnPrintTinNo = CBool(Find_Value("Select isnull(PrintTinNO,0) as PrintTinNO from sales_parameter where Unit_Code='" & gstrUnitId & "'"))
                If blnPrintTinNo = True Then
                    CR.DataDefinition.FormulaFields("TinNo").Text = "'" & TinNo & "'"
                End If

                SuppDeb_Print = CBool(Find_Value("Select isnull(Supp_DebNotePrinting,0) as SuppDeb_Print from customer_mst where Unit_Code='" & gstrUnitId & "' and Customer_code ='" & strCustomerCode & "' "))
                If SuppDeb_Print = True Then
                    CR.DataDefinition.FormulaFields("SuppDeb_Print").Text = " '1' "
                Else
                    CR.DataDefinition.FormulaFields("SuppDeb_Print").Text = " '0' "
                End If
                .ShowPrintButton = True
                .ShowExportButton = True
                .ShowTextSearchButton = True
                .WindowState = FormWindowState.Maximized
                CR.RecordSelectionFormula = strsql
            End With
        End If

        InvoiceGeneration = True
        Exit Function
Err_Handler:
        MsgBox(Err.Number)

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub dtFromDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtFromDate.ValueChanged
        If dtFromDate.Value > dtToDate.Value Then dtFromDate.Value = dtToDate.Value
    End Sub

    Private Sub dtToDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtToDate.ValueChanged
        If dtToDate.Value < dtFromDate.Value Then dtToDate.Value = dtFromDate.Value
    End Sub

    Private Sub optDrnote_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optDrnote.CheckedChanged
        Try
            refreshform()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub optCrnote_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCrnote.CheckedChanged
        Try
            refreshform()
            If optCrnote.Checked = True Then
                optCustomerCopy.Enabled = False
                OptOriginalcopy.Checked = True
            Else
                optCustomerCopy.Enabled = True
                OptOriginalcopy.Checked = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub OptGenerate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptGenerate.CheckedChanged
        If OptGenerate.Checked = True Then
            btn_print.Enabled = False
        Else
            btn_print.Enabled = True
        End If
    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

#Region "TO CREATE AND SAVE THE PDF AFTER LOCKING BY SUMIT KUMAR ON 21082019"
    Private Sub btnSavePDFBinary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSavePDFBinary.Click
        Try
            Dim blnnoitemselected As Boolean = False
            Dim intmainloop As Integer = 0
            If OptReprint.Checked AndAlso btnSavePDFBinary.Visible = True Then
                If txtFromInvoice.Text = "" Then
                    MsgBox("Please select invoice no (From Invoice No) !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    txtFromInvoice.Focus()
                    Exit Sub
                End If
                If txtToInvoice.Text = "" Then
                    MsgBox("Please Select invoice no (To Invoice No ) !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    txtToInvoice.Focus()
                    Exit Sub
                End If
                If dtFromDate.Value > dtToDate.Value Then
                    MsgBox("Invalid date range !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Exit Sub
                End If
                With fsprDtls
                    For intmainloop = 1 To .MaxRows
                        .Row = intmainloop
                        .Col = ENUMGRIDDTLS.ENUM_Mark

                        If CBool(.Value) = True Then
                            blnnoitemselected = True
                            Exit For
                        End If
                    Next
                    If Not blnnoitemselected Then
                        MsgBox("Please Select atleast one invoice !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        .Row = 1
                        .Col = ENUMGRIDDTLS.ENUM_Mark
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                        Exit Sub
                    End If
                End With
            End If
            GETDATA_INVOICELOCK_PDF()
            CREATEANDSAVE_SUPPLEMENTARYPDFINOIVE_AFTERLOCKING()
            lblpdfcreate.Visible = False
            FLAGISCUSTOMER_PDF_REQ = False
            If OptReprint.Checked AndAlso btnSavePDFBinary.Visible Then
                btnSavePDFBinary.Visible = False
                txtFromInvoice.Text = ""
                txtToInvoice.Text = ""
                fsprDtls.MaxRows = 0
            End If
        Catch ex As Exception
            FLAGISCUSTOMER_PDF_REQ = False
            lblpdfcreate.Text = ex.ToString()
        End Try
    End Sub
    Private Sub GETDATA_INVOICELOCK_PDF()
        Dim sqlCmd As New SqlCommand()
        Dim SqlAdp As New SqlDataAdapter
        Dim DSFILEDTL As New DataSet
        Dim strsql As String = String.Empty
        Dim intLoopCounter As Int32 = 0
        Dim STRDRCR As String
        Try

            With sqlCmd
                .CommandText = "USP_MULTI_SUPP_INV_DETAILS"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "R")
                If optDrnote.Checked = True Then
                    STRDRCR = "DR"
                Else
                    STRDRCR = "CR"
                End If
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@FROM_DATE", getDateForDB(dtFromDate.Value))
                .Parameters.AddWithValue("@TO_DATE", getDateForDB(dtToDate.Value))
                .Parameters.AddWithValue("@FROM_INVOICE", txtFromInvoice.Text.Trim)
                .Parameters.AddWithValue("@TO_INVOICE", txtToInvoice.Text.Trim)
                .Parameters.AddWithValue("@DRCR", STRDRCR)
                .Parameters.AddWithValue("@EWAY_BILL_FUNCTIONALITY", IIf(mblnEwayFunctionality = True, 1, 0))
                '.Parameters.AddWithValue("@ERR" 
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSFILEDTL)
                .Dispose()
            End With

            If DSFILEDTL.Tables.Count > 0 Then
                'DTL DATA
                If DSFILEDTL.Tables(0).Rows.Count > 0 Then
                    InitializeSpread_FileDtls()
                    With Me.fsprDtls
                        For intLoopCounter = 0 To DSFILEDTL.Tables(0).Rows.Count - 1
                            AddRow()
                            .SetText(ENUMGRIDDTLS.ENUM_TEMP_INVNO, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("TEMP_INVNO").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_CUSTOMERCODE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("CUSTOMER_CODE").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_CUSTNAME, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("CUST_NAME").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_ITEMCODE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("ITEM_CODE").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_RATEDIFF, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("RATE_DIFF").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_QTY, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("QUANTITY").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_TAXABE_AMT, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("TAXABLE_AMT").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_VALUE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("AMOUNT").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_TOTAL_INVAMT, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("TOTAL_AMOUNT").ToString.Trim)
                            .SetText(ENUMGRIDDTLS.ENUM_AR_DEBITNOTE, intLoopCounter + 1, DSFILEDTL.Tables(0).Rows(intLoopCounter).Item("VOUCHER_NO").ToString.Trim)
                        Next
                    End With
                Else
                    fsprDtls.MaxRows = 0
                    fsprDtls.MaxCols = 0
                    MsgBox("No Data Found For the Date Range Selected!", MsgBoxStyle.Information, "Empro")
                    txtFromInvoice.Text = ""
                    txtToInvoice.Text = ""
                End If
            Else
                MsgBox("No Data Found For the Date Range Selected!", MsgBoxStyle.Information, "Empro")
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            DSFILEDTL.Clear()
            DSFILEDTL.Dispose()
        End Try

    End Sub
    Private Sub CREATEANDSAVE_SUPPLEMENTARYPDFINOIVE_AFTERLOCKING()
        ''SUMIT KUMAR ON 21 AUG 2019

        Dim minvno As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intNoCopies As Short
        Dim intmainloop As Short
        Dim strinvoiceno As String = ""
        Dim strprevinvoiceno As String = ""
        Dim BLNselectanyinvoice As String = False
        Dim strsql As String
        Dim strvouchernumber As Object
        Dim strRepPath As String
        Dim COPYNAME As String
        Dim strPath As String = String.Empty
        Dim strinvPath As String = String.Empty


        With fsprDtls
            If .MaxRows > 0 Then
                For lngCounter = 1 To .MaxRows
                    .Row = lngCounter : .Col = ENUMGRIDDTLS.ENUM_Mark : .Value = CStr(1)
                Next
            End If
        End With


        frmRpt = New eMProCrystalReportViewer
        CR = New ReportDocument
        CR = frmRpt.GetReportDocument

        With fsprDtls
            If .MaxRows <= 0 Then

                Exit Sub
            End If
            For intmainloop = 1 To .MaxRows

                .Row = intmainloop
                .Col = ENUMGRIDDTLS.ENUM_Mark
                If CBool(.Value) = True Then
                    strinvoiceno = ""
                    .Row = intmainloop
                    .Col = ENUMGRIDDTLS.ENUM_TEMP_INVNO
                    strinvoiceno = .Text

                    .Col = ENUMGRIDDTLS.ENUM_AR_DEBITNOTE
                    strvouchernumber = Nothing
                    strvouchernumber = .Text
                    strcustomercode = Nothing
                    .Col = ENUMGRIDDTLS.ENUM_CUSTOMERCODE
                    strcustomercode = .Text

                    Dim OBJPdfConfig As Object = SqlConnectionclass.ExecuteScalar("SELECT COUNT(*) FROM INVOICE_PDF_CONFIG (NOLOCK) WHERE UNIT_CODE='" + gstrUnitId + "' AND CUSTOMER_CODE='" + strcustomercode.ToString() + "' AND INVOICE_TYPE='SUP' AND INVOICE_SUB_TYPE='S' And IS_ACTIVE=1")
                    If Not Val(OBJPdfConfig.ToString()) > 0 Then
                        GoTo again
                    End If
                    '' strPath = Find_Value("select Invoice_PDFCOPYPATH from CUSTOMER_MST  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strcustomercode & "'")

                    If strprevinvoiceno <> strinvoiceno Then
                        If File.Exists(strinvPath) = True Then
                            Kill(strinvPath)
                        End If
                    End If
                    mintnocopies = CInt(Find_Value("select ISNULL(NOOFCOPIES_SUPPINVOICE,0) NOOFCOPIES_SUPPINVOICE from CUSTOMER_MST  WHERE UNIT_CODE='" + gstrUnitId + "'AND CUSTOMER_CODE='" & strcustomercode & "'"))

                    If mblnEwayFunctionality Then
                        Call IRN_QRBarcode(strvouchernumber)
                    End If

                    If optDrnote.Checked = True Then
                        If intmainloop = 1 Or strprevinvoiceno <> strinvoiceno Then
                            strprevinvoiceno = strinvoiceno
                            BLNselectanyinvoice = True
                            If InvoiceGeneration(strinvoiceno) = True Then
                                intMaxLoop = mintnocopies
                                For intLoopCounter = 1 To intMaxLoop
                                    If OptReprint.Checked = True Then
                                        COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUnitId + "' AND CUSTOMER_CODE='" & strcustomercode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                    Else
                                        COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUnitId + "' AND CUSTOMER_CODE='" & strcustomercode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                    End If
                                    CR.DataDefinition.FormulaFields("CopyName").Text = "'" & COPYNAME & "'"
                                    frmRpt.SetReportDocument()

                                    'CR.ExportToDisk(ExportFormatType.PortableDocFormat, strinvPath)
                                    EXPORTINVOICETOPDF_ONPRINTREPRINT(strcustomercode, "SUP", "S", CR, COPYNAME, strvouchernumber)

                                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                                Next

                            End If
                        End If
                    Else
                        If DataExecution(strcustomercode, strvouchernumber) = False Then
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                            Exit Sub
                        End If

                        mstrReportFilename = "AR_Note_GST"
                        'strRepPath = My.Application.Info.DirectoryPath & strReportName
                        strRepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt"
                        'CR.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")

                        CR = New ReportDocument()
                        frmRpt = New eMProCrystalReportViewer()
                        CR = frmRpt.GetReportDocument()
                        CR.Load(strRepPath)
                        With CR
                            .SetParameterValue("repmaintitle", "Receipt voucher (Illustrative)")
                            .SetParameterValue("cname", gstrCOMPANY)
                            .SetParameterValue("address1", Trim(gstr_WRK_ADDRESS1))
                            .SetParameterValue("address2", Trim(gstr_WRK_ADDRESS2))

                            .RecordSelectionFormula = "{TMP_FIN_GET_GST_AR_NOTE_DETAIL.IP_ADDRESS}='" & gstrIpaddressWinSck & "'  and {TMP_FIN_GET_GST_AR_NOTE_DETAIL.DOCM_UNIT}='" & gstrUnitId & "'"
                        End With
                        'frmRpt.Show()
                        frmRpt.SetReportDocument()

                        BLNselectanyinvoice = True
                        If intmainloop = 1 Or strprevinvoiceno <> strinvoiceno Then
                            strprevinvoiceno = strinvoiceno
                            EXPORTINVOICETOPDF_ONPRINTREPRINT(strcustomercode, "SUP", "S", CR, COPYNAME, strvouchernumber)

                        End If

                    End If
                End If
again:
            Next


        End With
    End Sub
    Private Sub EXPORTINVOICETOPDF_ONPRINTREPRINT(ByVal strAccountCode As String, ByVal strInvoiceType As String, ByVal strInvoiceSubType As String, ByRef RPTDoc As ReportDocument, ByVal COPYRPT As String, ByVal strInvoiceno As String)
        Dim strCreatedPDFPath As String = String.Empty
        Dim strRESULT As Collections.ArrayList

        RPTDoc.Export(GetExportOptions(strInvoiceno, strCreatedPDFPath))
        Dim OBJPdfConfig As Object = SqlConnectionclass.ExecuteScalar("SELECT COUNT(*) FROM INVOICE_PDF_CONFIG (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" + strAccountCode.ToString() + "' AND INVOICE_TYPE='SUP' AND INVOICE_SUB_TYPE='S' And IS_ACTIVE=1")
        Dim OBJCommonDigital_EINVOICING_CONFIG As Object = SqlConnectionclass.ExecuteScalar("select Count(*) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" + gstrUNITID + "' and CUSTOMER_CODE='" + strAccountCode.ToString() + "' AND INVOICE_TYPE='SUP' and SUB_TYPE='S' AND IS_ACTIVE=1")
        strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), strInvoiceno, "SUPPLEMENTARY", gstrUNITID, strCreatedPDFPath, "SUP", "S", OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG, 0)
        If strRESULT.Item(0).ToString().Trim().ToUpper() <> "SUCCESS" Then
            If strRESULT.Item(0).ToString().Trim().ToUpper() <> "FAIL" Then
                MsgBox(strRESULT.Item(0).ToString().Trim())
            End If
        End If



    End Sub
    Private Function GetExportOptions(ByVal strInvoiceNoForFileName As String, ByRef strCreatedPDFPath As String) As ExportOptions
        Dim strPath As String
        If strCreatedPDFPath = "" Then
            strPath = Find_Value("select Invoice_PDFCOPYPATH from CUSTOMER_MST  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strcustomercode & "'")
            If (System.IO.Directory.Exists(strPath) = False) Then
                System.IO.Directory.CreateDirectory(strPath)
            End If
            strCreatedPDFPath = strPath + "\" + strInvoiceNoForFileName + ".pdf"
        End If

        Dim fileDestinationOptions As New DiskFileDestinationOptions
        Dim exportOptions As New ExportOptions()
        fileDestinationOptions.DiskFileName = strCreatedPDFPath 'eInvoicingFileName
        exportOptions.ExportDestinationOptions = fileDestinationOptions
        exportOptions.ExportDestinationType = ExportDestinationType.DiskFile
        exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
        Return exportOptions
    End Function


    Private Sub IRN_QRBarcode(ByVal voucherNo As String)
        Try
            Dim rsGENERATEBARCODE As ClsResultSetDB
            Dim straccountcode As String
            Dim strPrintMethod As String = ""
            Dim strSQL As String = ""
            Dim intTotalNoofSlabs As Integer = 0
            Dim intRow As Short
            Dim strBarcodeMsg As String
            Dim strBarcodeMsg_paratemeter As String
            Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUnitId)
            Dim stimage As ADODB.Stream
            Dim strQuery As String
            Dim Rs As ADODB.Recordset
            Dim pstrPath As String = ""
            Dim blnCROP_QRIMAGE As Boolean = False


            pstrPath = gstrUserMyDocPath
            strSQL = "SELECT TOP 1 1 FROM Supplementary_IRN I INNER JOIN Supplementary_IRN_BARCODE B ON I.UNIT_CODE=B.UNIT_CODE AND I.VO_NO=B.VO_NO WHERE I.UNIT_CODE = '" & gstrUnitId & "' AND I.VO_NO='" & Trim(voucherNo) & "'" & " AND ISNULL(I.IRN_NO,'')<>'' AND ISNULL(B.BARCODE_DATA,'')<>'' "

            If DataExist(strSQL) = True Then
                strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForSupplementaryIRN(gstrUserMyDocPath, Trim(voucherNo), gstrCONNECTIONSTRING)

                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                    Exit Sub
                Else
                    strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                    stimage = New ADODB.Stream
                    stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
                    stimage.Open()
                    pstrPath = pstrPath & "QRBarcodeImgSuppIRN.wmf"

                    blnCROP_QRIMAGE = CBool(Find_Value("SELECT CROP_IRN_QRBARCODE  FROM SALES_PARAMETER (NOLOCK) WHERE UNIT_CODE='" + gstrUnitId + "'"))
                    If blnCROP_QRIMAGE = True Then
                        Dim bmp As New Bitmap(pstrPath)
                        Dim picturebox1 As New PictureBox
                        picturebox1.Image = ImageTrim(bmp)
                        picturebox1.Image.Save(pstrPath)
                        picturebox1 = Nothing
                    End If

                    stimage.LoadFromFile(pstrPath)

                    strQuery = "select  BARCODE_DATA,VO_NO ,barcodeimage from Supplementary_IRN_BARCODE where UNIT_CODE = '" & gstrUnitId & "' AND VO_No='" & Trim(voucherNo) & "' "

                    Rs = New ADODB.Recordset
                    Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                    If Not (Rs.EOF And Rs.BOF) Then
                        Rs.Fields("barcodeimage").Value = stimage.Read
                        Rs.Update()
                    End If

                    Rs.Update()
                    Rs.Close()
                    Rs = Nothing


                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Function ImageTrim(ByVal img As Bitmap) As Bitmap
        'get image data
        Dim bd As BitmapData = img.LockBits(New Rectangle(Point.Empty, img.Size), ImageLockMode.[ReadOnly], PixelFormat.Format32bppArgb)
        Dim rgbValues As Integer() = New Integer(img.Height * img.Width - 1) {}
        Marshal.Copy(bd.Scan0, rgbValues, 0, rgbValues.Length)
        img.UnlockBits(bd)


        '#Region "determine bounds"
        Dim left As Integer = bd.Width
        Dim top As Integer = bd.Height
        Dim right As Integer = 0
        Dim bottom As Integer = 0

        'determine top
        For i As Integer = 0 To rgbValues.Length - 1
            Dim color As Integer = rgbValues(i) And &HFFFFFF
            If color <> &HFFFFFF Then
                Dim r As Integer = i / bd.Width
                Dim c As Integer = i Mod bd.Width

                If left > c Then
                    left = c
                End If
                If right < c Then
                    right = c
                End If
                bottom = r
                top = r
                Exit For
            End If
        Next

        'determine bottom
        For i As Integer = rgbValues.Length - 1 To 0 Step -1
            Dim color As Integer = rgbValues(i) And &HFFFFFF
            If color <> &HFFFFFF Then
                Dim r As Integer = i / bd.Width
                Dim c As Integer = i Mod bd.Width

                If left > c Then
                    left = c
                End If
                If right < c Then
                    right = c
                End If
                bottom = r
                Exit For
            End If
        Next

        If bottom > top Then
            For r As Integer = top + 1 To bottom - 1
                'determine left
                For c As Integer = 0 To left - 1
                    Dim color As Integer = rgbValues(r * bd.Width + c) And &HFFFFFF
                    If color <> &HFFFFFF Then
                        If left > c Then
                            left = c
                            Exit For
                        End If
                    End If
                Next

                'determine right
                For c As Integer = bd.Width - 1 To right + 1 Step -1
                    Dim color As Integer = rgbValues(r * bd.Width + c) And &HFFFFFF
                    If color <> &HFFFFFF Then
                        If right < c Then
                            right = c
                            Exit For
                        End If
                    End If
                Next
            Next
        End If

        Dim width As Integer = right - left + 1
        Dim height As Integer = bottom - top + 1
        '#End Region

        'copy image data
        Dim imgData As Integer() = New Integer(width * height - 1) {}
        For r As Integer = top To bottom
            Array.Copy(rgbValues, r * bd.Width + left, imgData, (r - top) * width, width)
        Next

        'create new image
        Dim newImage As New Bitmap(width, height, PixelFormat.Format32bppArgb)
        Dim nbd As BitmapData = newImage.LockBits(New Rectangle(0, 0, width, height), ImageLockMode.[WriteOnly], PixelFormat.Format32bppArgb)
        Marshal.Copy(imgData, 0, nbd.Scan0, imgData.Length)
        newImage.UnlockBits(nbd)

        ImageTrim = newImage
    End Function
#End Region
End Class