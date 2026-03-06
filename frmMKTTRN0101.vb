Option Strict Off
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports VB = Microsoft.VisualBasic
Imports Excel = Microsoft.Office.Interop.Excel
Friend Class frmMKTTRN0101
    Inherits System.Windows.Forms.Form
    '****************************************************
    'Copyright (c)  -  MIND
    'Name of module -  FRMMKTTRN0066.frm
    'Created By     -  Prashant Rajpal
    'Created On     -  22 jan 2018
    '****************************************************
    Dim blnISBasicRoundOff As Boolean
    Dim intBasicRoundOffDecimal As Integer
    Dim blnTotalInvoiceAmount As Boolean
    Dim intTotalInvoiceAmountRoundOffDecimal As Integer
    Dim intGSTTAXroundoff_decimal As Short
    Dim blnGSTTAXroundoff As Boolean
    Dim m_strCustomerCode As String
    Dim mintFormIndex As Short
    Dim StrDocNum As String
    Dim mobjEmpDll As New EMPDataBase.EMPDB(gstrUNITID)
    Dim mrsEmpDll As New EMPDataBase.CRecordset
    Dim mbln_View As Boolean
    Dim Flag As Boolean
    Dim dtSelItems As DataTable
    Dim dtDocTable As DataTable
    Dim dblBasicValue As Decimal
    Dim DocFrm As Form
    Dim part_code As String = String.Empty
    Dim rate As Double = 0.0
    Dim strPartNo As String = String.Empty
    Dim stritemcode As String = String.Empty
    Dim strpartdesc As String = String.Empty
    Dim strHSNSACCODE As String = String.Empty
    Dim strBillNo As String
    Dim strPricDt As String = String.Empty
    Dim strBilldate As String = String.Empty
    Dim strshp As Int16 = 0
    Dim stracp As Int16 = 0
    Dim dbloldrate As Decimal = 0.0
    Dim dblNewRate As Decimal = 0.0
    Dim strcgstamt As Decimal = 0.0
    Dim strsgstamt As Decimal = 0.0
    Dim strigstamt As Decimal = 0.0
    Dim strsrvdino As String
    Dim strbasicamt As Decimal = 0.0
    Dim dbldiffrate As Decimal = 0.0
    Dim dbltotalamt As Double
    Dim str57f2 As String
    Dim blnSUPP_INVOICE_ONEITEM_MULTIPLE As Boolean
    Dim i As Integer = 0
    Dim mblnBatchFile_NewFomrat As Boolean = False

    Private Enum ENUM_GRID
        INVOICENO = 1
        DOC_NO
        ITEMCODE
        PARTNO
        DESCRIPTION
        BILLNO
        SHIPQTY
        ACTQTY
        RATEDIFF
        HSNSACCODE
        BASICAMT
        CGSTTAX
        CGSTAMT
        SGSTTAX
        SGSTAMT
        IGSTTAX
        IGSTAMT
        SRVIDINO
        BATCHNO
        SERIAL_NO

    End Enum
    Private Sub FN_Spread_Settings()
        Dim Col As Integer
        Try
            With SSSUPPLEMENTARYDATA
                .MaxCols = ENUM_GRID.SERIAL_NO
                .MaxRows = 0
                .Row = 0
                .Col = ENUM_GRID.INVOICENO : .Text = "BASE INV NO"
                .Col = ENUM_GRID.DOC_NO : .Text = "INV. NO" : .ColHidden = True
                .Col = ENUM_GRID.ITEMCODE : .Text = "ITEM CODE"
                .Col = ENUM_GRID.PARTNO : .Text = "PART NO"
                .Col = ENUM_GRID.DESCRIPTION : .Text = "PART DESC"
                .Col = ENUM_GRID.BILLNO : .Text = "BILL NO"
                .Col = ENUM_GRID.SHIPQTY : .Text = "SHIP QTY"
                .Col = ENUM_GRID.ACTQTY : .Text = "INVOICE QTY"
                .Col = ENUM_GRID.RATEDIFF : .Text = "INVOICE RATE DIFF"
                .Col = ENUM_GRID.BASICAMT : .Text = "TAXABLE AMOUNT"
                .Col = ENUM_GRID.HSNSACCODE : .Text = "HSN/SAC"
                .Col = ENUM_GRID.CGSTTAX : .Text = "CGST TAX"
                .Col = ENUM_GRID.CGSTAMT : .Text = "CGST AMT"
                .Col = ENUM_GRID.SGSTTAX : .Text = "SGST TAX"
                .Col = ENUM_GRID.SGSTAMT : .Text = "SGST AMT"
                .Col = ENUM_GRID.IGSTTAX : .Text = "IGST TAX"
                .Col = ENUM_GRID.IGSTAMT : .Text = "IGST AMT"
                .Col = ENUM_GRID.SRVIDINO : .Text = "MARUTI SRY. NO"
                .Col = ENUM_GRID.BATCHNO : .Text = "BATCH NO"
                .Col = ENUM_GRID.SERIAL_NO : .Text = "SERIAL NO"
                .set_RowHeight(0, 20)
                .set_ColWidth(ENUM_GRID.INVOICENO, 10)
                .set_ColWidth(ENUM_GRID.ITEMCODE, 12)
                .set_ColWidth(ENUM_GRID.PARTNO, 12)
                .set_ColWidth(ENUM_GRID.DESCRIPTION, 30)
                .set_ColWidth(ENUM_GRID.BILLNO, 0)
                .set_ColWidth(ENUM_GRID.SHIPQTY, 0)
                .set_ColWidth(ENUM_GRID.ACTQTY, 6)
                .set_ColWidth(ENUM_GRID.RATEDIFF, 6)
                .set_ColWidth(ENUM_GRID.BASICAMT, 8)
                .set_ColWidth(ENUM_GRID.HSNSACCODE, 8)
                .set_ColWidth(ENUM_GRID.CGSTTAX, 8)
                .set_ColWidth(ENUM_GRID.SGSTTAX, 8)
                .set_ColWidth(ENUM_GRID.IGSTTAX, 8)
                .set_ColWidth(ENUM_GRID.CGSTAMT, 8)
                .set_ColWidth(ENUM_GRID.SGSTAMT, 8)
                .set_ColWidth(ENUM_GRID.IGSTAMT, 8)
                .set_ColWidth(ENUM_GRID.SRVIDINO, 10)
                .set_ColWidth(ENUM_GRID.BATCHNO, 8)
                .set_ColWidth(ENUM_GRID.SERIAL_NO, 10)

                .Row = 1
                .Row = .MaxRows
                .Col = ENUM_GRID.INVOICENO
                .Col2 = .MaxCols
                .Lock = True
                .BlockMode = True
                .Width = 85 * 8.5
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustHelp.Click
        Dim strHelp() As String
        Try
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select customer_Code,cust_name from customer_mst where UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", "Help", 2)
            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    If strHelp(0) <> "" Then
                        txtCustomerCode.Text = strHelp(0)
                        LblCustomerName.Text = strHelp(1)
                        txtDocNo.Text = ""
                        cmdUpload.Enabled = True
                        cmdSave.Enabled = True
                        txtCustomerCode.Focus()
                    End If
                Else
                    Me.txtCustomerCode.Text = ""
                    Me.LblCustomerName.Text = ""
                    Me.txtCustomerCode.Focus()
                    MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                End If
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub cmdFileHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFileHelp.Click
        Try
            CommanDLogOpen.InitialDirectory = "d:\"
            CommanDLogOpen.Filter = "Microsoft text File (*.txt)|*.txt"
            CommanDLogOpen.ShowDialog()
            Me.txtFileName.Text = CommanDLogOpen.FileName
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub cmdHelpDocNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpDocNo.Click
        Dim StrSql As String
        Dim docno As Integer
        Dim strHelp() As String
        Try
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            StrSql = "SELECT DISTINCT CAST(M.DOCNO AS VARCHAR(10)) AS DOC_NO , CONVERT(VARCHAR(20),ENT_DT,106) AS UPLOADDATE FROM MARUTI_SUPPINVOICE_FILE_HDR M WHERE M.UNIT_CODE='" & gstrUNITID & "'"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql)

            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    If strHelp(0) = "" Then
                        Exit Sub
                    End If
                    Me.txtDocNo.Text = strHelp(0)
                    docno = CInt(txtDocNo.Text)
                    Call txtDocNo_Validating(txtDocNo, New System.ComponentModel.CancelEventArgs(True))
                    mbln_View = True
                    DISPLAY()
                    cmdUpload.Enabled = False
                    cmdSave.Enabled = False
                Else
                    MsgBox(" No record available!", MsgBoxStyle.Information, ResolveResString(100))
                    mbln_View = False
                    cmdUpload.Enabled = True
                    cmdSave.Enabled = False
                End If
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub ctlFormHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader.Click
        Try
            Call ShowHelp("UNDERCONSTRUCTION.HTM")
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub frmMKTTRN0058_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try
            'Form number assigned for the current form
            mdifrmMain.CheckFormName = mintFormIndex
            'Form name text is made BOLD
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub frmMKTTRN0058_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Try
            'Form name would be adjusted to normal
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub frmMKTTRN0058_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
                Call ctlFormHeader_ClickEvent(ctlFormHeader, New System.EventArgs())
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub frmMKTTRN0058_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            'Adjusts the form to the main Window standards.
            Call FitToClient(Me, frmMain, ctlFormHeader, Panel1, 500)
            'Call FillLabelFromResFile(Me)

            mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.HeaderString())
            Call EnableControls(True, Me)
            Call FN_Spread_Settings()
            mblnBatchFile_NewFomrat = Find_Value("SELECT isnull(BatchFile_NewFomrat, 0) FROM sales_parameter  (Nolock) WHERE UNIT_CODE='" & gstrUNITID & "'")
            If mblnBatchFile_NewFomrat = True Then
                chkNewFormat.Visible = True
            Else
                chkNewFormat.Visible = False
            End If

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub frmMKTTRN0058_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
            frmModules.NodeFontBold(Me.Tag) = False
            Me.Dispose()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            LblCustomerName.Text = ""
            SSSUPPLEMENTARYDATA.MaxRows = 0
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdCustHelp_Click(cmdCustHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
            If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim bolExist As Boolean
        Try
            If Len(txtCustomerCode.Text) > 0 Then
                bolExist = ValCustomerCode()
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            eventArgs.Cancel = Cancel
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Public Function ValCustomerCode() As Boolean ' Checks for Customer Code
        Dim ms As String = ""
        Try
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            Call mobjEmpDll.CRecordset.OpenRecordset("select * from customer_mst where customer_code = '" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            mobjEmpDll.CRecordset.Filter_Renamed = "customer_code='" & Trim(txtCustomerCode.Text) & "'"
            If mobjEmpDll.CRecordset.Recordcount > 0 Then ValCustomerCode = True Else ValCustomerCode = False
            mobjEmpDll.CConnection.CloseConnection()
        Catch ex As Exception
            mobjEmpDll.CConnection.CloseConnection()
            mrsEmpDll.CloseRecordset()
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub txtDocNo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDocNo.KeyUp
        Try
            If txtDocNo.Text.Length = 0 Then
                txtCustomerCode.Text = ""
                'txtFileName.Text = ""
                SSSUPPLEMENTARYDATA.MaxRows = 0
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtDocNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDocNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim docno As Integer
        Dim StrSql As String
        Dim rsDocNo As New ClsResultSetDB
        Try
            If Len(txtDocNo.Text.Trim) > 0 Then
                docno = CInt(txtDocNo.Text)
                txtDocNo.Text = CStr(docno)
                StrSql = "SELECT H.DOCNO FROM MARUTI_SUPPINVOICE_FILE_HDR H WHERE H.DOCNO = '" & docno & "' and H.UNIT_CODE='" & gstrUNITID & "'"
                rsDocNo.GetResult(StrSql)
                If rsDocNo.GetNoRows > 0 Then
                    mbln_View = True
                Else
                    SSSUPPLEMENTARYDATA.MaxRows = 0
                    txtCustomerCode.Text = ""
                    txtFileName.Text = ""
                    MsgBox(" No record available!", MsgBoxStyle.Information, ResolveResString(100))
                    mbln_View = False
                    cmdUpload.Enabled = True
                End If
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            eventArgs.Cancel = Cancel
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtFileName_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFileName.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        Try
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdFileHelp_Click(cmdFileHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtFileName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFileName.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
            If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtDocNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDocNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        Try
            If (KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0) Then
                Call cmdHelpDocNo_Click(cmdHelpDocNo, New System.EventArgs())
            End If
            If (KeyCode = System.Windows.Forms.Keys.Return And Shift = 0) Then
                Call txtDocNo_Validating(txtDocNo, New System.ComponentModel.CancelEventArgs(True))
                If mbln_View = True Then
                    DISPLAY()
                    cmdUpload.Enabled = False
                    cmdSave.Enabled = False
                End If
            End If
            If txtDocNo.Text.Length = 0 Then
                txtCustomerCode.Text = ""
                txtFileName.Text = ""
                SSSUPPLEMENTARYDATA.MaxRows = 0
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtDocNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDocNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
            If KeyAscii = 8 Then
                Exit Sub
            End If
            If KeyAscii < 48 Or KeyAscii > 57 Then
                If KeyAscii <> 13 Then
                    KeyAscii = 0
                End If
            End If
            If KeyAscii = 13 Then
                System.Windows.Forms.SendKeys.Send("{tab}")
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub SubGetRoundoffConfig()
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim DataRd As SqlDataReader
        Try

            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text

            Sqlcmd.CommandText = "SELECT isnull(Basic_Roundoff,0) as Basic_Roundoff, isnull(Basic_Roundoff_decimal,0) as Basic_Roundoff_decimal, " &
            "isnull(TotalInvoiceAmount_RoundOff,0) as TotalInvoiceAmount_RoundOff,isnull(TotalInvoiceAmountRoundOff_Decimal,0) as TotalInvoiceAmountRoundOff_Decimal, " &
            "isnull(GSTTAX_ROUNDOFF,0) as GSTTAX_ROUNDOFF,isnull(GSTTAX_ROUNDOFF_DECIMAL,0) as GSTTAX_ROUNDOFF_DECIMAL , SUPP_INVOICE_ONEITEM_MULTIPLE FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
            DataRd = Sqlcmd.ExecuteReader()
            If DataRd.HasRows Then
                DataRd.Read()
                blnISBasicRoundOff = DataRd("Basic_Roundoff")
                intBasicRoundOffDecimal = DataRd("Basic_Roundoff_decimal")
                blnTotalInvoiceAmount = DataRd("TotalInvoiceAmount_RoundOff")
                intTotalInvoiceAmountRoundOffDecimal = DataRd("TotalInvoiceAmountRoundOff_Decimal")
                blnGSTTAXroundoff = DataRd("GSTTAX_ROUNDOFF")
                intGSTTAXroundoff_decimal = DataRd("GSTTAX_ROUNDOFF_DECIMAL")
                blnSUPP_INVOICE_ONEITEM_MULTIPLE = DataRd("SUPP_INVOICE_ONEITEM_MULTIPLE")
            Else
                MessageBox.Show("No Data Define In Sales_Parameter Table", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            If DataRd.IsClosed = False Then DataRd.Close()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsNothing(Sqlcmd.Connection) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                Sqlcmd.Connection.Dispose()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Dispose()
            End If
        End Try

    End Sub
    Private Function MUL_SCHEDULE_UPLOAD() As Object
        Try

            Dim Obj_EX As New Excel.Application
            Dim Row_Ex As Short
            Dim Col_Ex As Short
            Dim STRSQL As String = String.Empty
            Dim STRLENGTH As String = String.Empty
            Dim BATCHNO As String = String.Empty
            Dim batch_no As String = String.Empty
            Dim batch As String = String.Empty
            Dim strDataLine As String
            Dim strVal() As String = String.Empty.Split(",")
            Dim serial_no As Integer = 0
            Dim line As String
            Dim strOriLine As String
            Dim sequence_no As String = String.Empty
            Dim strpartcode As String = String.Empty
            Dim intRows As Short
            Dim strRow As String
            Dim intRowSeperator As Short
            Dim stcgsttaxtype As String
            Dim stsgsttaxtype As String
            Dim stcgsttaxper As String
            Dim stsgsttaxper As String
            Dim stigsttaxtype As String
            Dim stigsttaxper As String
            Dim strcgsttaxtype As String
            Dim strsgsttaxtype As String
            Dim strigsttaxtype As String
            Dim strpartcode1 As String
            Dim FileName As String = String.Empty

            If txtCustomerCode.Text = "" Then
                MsgBox("ENTER CUSTOMER CODE", MsgBoxStyle.Information, ResolveResString(100))
                Return False
                Exit Function
            End If

            If txtFileName.Text = "" Then
                MsgBox("ENTER FILE LOCATION", MsgBoxStyle.Information, ResolveResString(100))
                Return False
                Exit Function
            End If

            FileName = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(FILENAME,'') FILENAME FROM MARUTI_SUPPINVOICE_FILE_HDR  WHERE UNIT_CODE ='" & gstrUNITID & "' AND FILENAME ='" & txtFileName.Text.Trim.ToString() & "'"))
            If Not IsNothing(FileName) Then
                If FileName = txtFileName.Text.Trim.ToString() Then
                    MsgBox("Selected File Already Uploaded!!")
                    Return False
                    Exit Function
                End If
            End If

            SubGetRoundoffConfig()
            Dim fd As OpenFileDialog = New OpenFileDialog()
            STRSQL = "delete from TMP_MarutisupplementaryData where Unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
            mP_Connection.Execute(STRSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            mP_Connection.BeginTrans()
            Using sr As StreamReader = New StreamReader(txtFileName.Text)

                ' Read and display lines from the file until the end of  
                ' the file is reached. q
                While (Not sr.EndOfStream)
                    line = sr.ReadLine()
                    strOriLine = line
                    batch_no = line.Replace("BATCH NO  :", "§§")
                    i = batch_no.IndexOf("§§")
                    If (i > 0) Then
                        batch = batch_no.Substring(i + 2, 6).Trim()
                    End If

                    sequence_no = line.Replace("SERIAL NO :", "§§§")
                    i = sequence_no.IndexOf("§§§")
                    If (i > 0) Then
                        'serial_no = Convert.ToInt32(sequence_no.Substring(i + 3, 4).Trim())
                        serial_no = Convert.ToInt32(sequence_no.Substring(i + 3).Trim())
                    End If

                    strpartcode = line.Replace("Part No. :", "§")


                    i = strpartcode.IndexOf("§")

                    If (i > 0) Then
                        strpartcode1 = strpartcode.Substring(i + 2, 40)
                        strPartNo = strpartcode.Substring(i + 2, strpartcode1.IndexOf(" "))
                        stritemcode = Find_Value("SELECT TOP 1 item_code FROM custitem_mst WHERE UNIT_CODE='" & gstrUNITID & "' AND cust_drgno='" & strPartNo & "' and account_code='" & txtCustomerCode.Text & "'") 'and active=1
                        strpartdesc = Find_Value("SELECT TOP 1 Drg_Desc FROM custitem_mst WHERE UNIT_CODE='" & gstrUNITID & "' AND cust_drgno='" & strPartNo & "' and account_code='" & txtCustomerCode.Text & "'") 'and active=1
                        strHSNSACCODE = Find_Value("SELECT TOP 1 HSN_SAC_CODE FROM item_mst WHERE UNIT_CODE='" & gstrUNITID & "' AND item_code ='" & stritemcode & "'")
                    End If

                    stcgsttaxtype = line.Replace("CGST :", "§§§")
                    i = stcgsttaxtype.IndexOf("§§§")
                    If (i > 0) Then
                        stcgsttaxtype = stcgsttaxtype.Substring(i + 5, 2)
                        stcgsttaxper = stcgsttaxtype
                        strcgsttaxtype = "CGST" + stcgsttaxper
                    End If

                    stsgsttaxtype = line.Replace("SGST :", "§§§")
                    i = stsgsttaxtype.IndexOf("§§§")
                    If (i > 0) Then
                        stsgsttaxtype = stsgsttaxtype.Substring(i + 5, 2)
                        stsgsttaxper = stsgsttaxtype
                        strsgsttaxtype = "SGST" + stsgsttaxper
                    End If

                    If (IsNumeric(stsgsttaxper) = False Or stsgsttaxper = "0") Then
                        If stigsttaxper = 0 Then
                            stigsttaxtype = line.Replace("IGST :", "§§§")
                            i = stigsttaxtype.IndexOf("§§§")
                            If (i > 0) Then
                                stigsttaxtype = stigsttaxtype.Substring(i + 5, 2)
                                stigsttaxper = stigsttaxtype
                                strigsttaxtype = "IGST" + stigsttaxper
                            End If
                        Else
                            If stigsttaxtype = "" Then
                                stigsttaxtype = "IGST0"
                                stigsttaxper = "0.00"
                            End If
                        End If
                    Else
                        stigsttaxtype = "IGST0"
                        stigsttaxper = "0.00"
                    End If

                    If (strVal.Length = 1) Then
                        line = line.Replace(" ", ",").Replace(",,,,,,,,,,", ",").Replace(",,,,,,,,,", ",").Replace(",,,,,,,,", ",").Replace(",,,,,,,", ",").Replace(",,,,,,", ",").Replace(",,,,,", ",").Replace(",,,,", ",").Replace(",,,", ",").Replace(",,", ",")


                        intRows = UBound(Split(line, ","))
                        'strVal = strDataLine.Split(",")

                        If intRowSeperator > 1 Then
                            strRow = Mid(line, 1, intRowSeperator - 1) 'Get 1 Row at a Time
                        End If
                        Dim intPos, counter As Integer
                        Dim strFields(17) As String
                        Dim strMstFields() As String

                        Dim strMaster As String
                        Dim inti As Integer

                        If intRows = 21 Then
                            counter = 0
                            intRowSeperator = InStr(line, ",")
                            strMaster = CStr(line).ToString
                            intPos = 0
                            counter = 20

                            While Len(strMaster) > 0
                                ReDim Preserve strMstFields(counter)
                                For inti = 0 To counter
                                    intPos = InStr(strMaster, ",")
                                    If intPos > 0 Then
                                        strMstFields(inti) = Mid(strMaster, 1, intPos - 1)
                                        strMaster = Mid(strMaster, intPos + 1)
                                    Else
                                        strMstFields(inti) = strMaster
                                        strMaster = ""
                                        Exit While
                                    End If
                                Next
                            End While

                            If IsDate(strMstFields(3)) = True Then
                                strBillNo = strMstFields(2).ToString
                                str57f2 = strMstFields(4).ToString
                                strPricDt = strMstFields(3).ToString
                                strshp = strMstFields(6).ToString
                                stracp = strMstFields(7).ToString
                                dbloldrate = Convert.ToDecimal(strMstFields(8))
                                strBilldate = strMstFields(14).ToString()
                                ' added by priti on 05.06.2019 to fix credit debit not equal
                                If blnISBasicRoundOff = True Then
                                    strbasicamt = Convert.ToDecimal(strMstFields(9))
                                Else
                                    strbasicamt = System.Math.Round(Convert.ToDecimal(strMstFields(9)), intBasicRoundOffDecimal)
                                End If
                                ' code ends here
                                If Trim(stcgsttaxper) = "%" Then
                                    stcgsttaxper = "0"
                                End If
                                If Trim(stsgsttaxper) = "%" Then
                                    stsgsttaxper = "0"
                                End If
                                ' added by priti on 05.06.2019 to fix credit debit not equal
                                If blnGSTTAXroundoff = True Then
                                    strcgstamt = strbasicamt * stcgsttaxper / 100
                                    strsgstamt = strbasicamt * stsgsttaxper / 100
                                    strigstamt = strbasicamt * stigsttaxper / 100
                                Else
                                    strcgstamt = System.Math.Round(strbasicamt * stcgsttaxper / 100, intGSTTAXroundoff_decimal)
                                    strsgstamt = System.Math.Round(strbasicamt * stsgsttaxper / 100, intGSTTAXroundoff_decimal)
                                    strigstamt = System.Math.Round(strbasicamt * stigsttaxper / 100, intGSTTAXroundoff_decimal)
                                End If
                                ' code ends here

                                dblNewRate = Convert.ToDecimal(strMstFields(18))
                                strsrvdino = strMstFields(5).ToString()
                                dbldiffrate = dblNewRate - dbloldrate
                                dbltotalamt = strbasicamt + strcgstamt + strsgstamt + strigstamt
                                If dbldiffrate > 0 Then
                                    STRSQL = "Insert Into TMP_MarutisupplementaryData (strPartNo,part_desc,item_code,strBillNo,strPricDt,strshp,stracp,stroldrate,strNewRate,DIFF_RATE,Unit_code,Ipaddress,invoiceNo,batch_no,serial_no,strBilldate,srvdino,Account_Code,basic_amount,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,hsnsaccode,TOTAL_AMOUNT ) Values ('" & strPartNo & "','" & strpartdesc & "','" & stritemcode & "','" & strBillNo & "','" & strPricDt & "','" & strshp & "','" & stracp & "','" & dbloldrate & "','" & dblNewRate & "','," & dbldiffrate & "','" & gstrUNITID & "','" & gstrIpaddressWinSck & "','" & str57f2 & "','" & batch & "'," & serial_no & ",'" & strBilldate & "','" & strsrvdino & "','" & txtCustomerCode.Text & "'," & strbasicamt & ",'" & strcgsttaxtype & "','" & stcgsttaxper & "'," & strcgstamt & ",'" & strsgsttaxtype & "','" & stsgsttaxper & "'," & strsgstamt & ",'" & strigsttaxtype & "','" & stigsttaxper & "'," & strigstamt & ",'" & strHSNSACCODE & "'," & dbltotalamt & " )"
                                    mP_Connection.Execute(STRSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If

                            STRSQL = ""
                        End If
                    End If
                End While
                STRSQL = ""
                mP_Connection.CommitTrans()
                'txtFileName.Text = ""

            End Using
            Return True
        Catch ex As Exception
            'RaiseException(ex)
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try

    End Function

    '' INC1557693 - Added by priti on 21 Nov 2025 for adding new format option.
    Private Function MUL_SCHEDULE_UPLOAD_NEWFORMAT() As Object
        Try

            Dim Obj_EX As New Excel.Application
            Dim Row_Ex As Short
            Dim Col_Ex As Short
            Dim STRSQL As String = String.Empty
            Dim STRLENGTH As String = String.Empty
            Dim BATCHNO As String = String.Empty
            Dim batch_no As String = String.Empty
            Dim batch As String = String.Empty
            Dim strDataLine As String
            Dim strVal() As String = String.Empty.Split(",")
            Dim serial_no As Integer = 0
            Dim line As String
            Dim strOriLine As String
            Dim sequence_no As String = String.Empty
            Dim strpartcode As String = String.Empty
            Dim intRows As Short
            Dim strRow As String
            Dim intRowSeperator As Short
            Dim stcgsttaxtype As String
            Dim stsgsttaxtype As String
            Dim stcgsttaxper As String
            Dim stsgsttaxper As String
            Dim stigsttaxtype As String
            Dim stigsttaxper As String
            Dim strcgsttaxtype As String
            Dim strsgsttaxtype As String
            Dim strigsttaxtype As String
            Dim strpartcode1 As String
            Dim FileName As String = String.Empty

            If txtCustomerCode.Text = "" Then
                MsgBox("ENTER CUSTOMER CODE", MsgBoxStyle.Information, ResolveResString(100))
                Return False
                Exit Function
            End If

            If txtFileName.Text = "" Then
                MsgBox("ENTER FILE LOCATION", MsgBoxStyle.Information, ResolveResString(100))
                Return False
                Exit Function
            End If

            FileName = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(FILENAME,'') FILENAME FROM MARUTI_SUPPINVOICE_FILE_HDR  WHERE UNIT_CODE ='" & gstrUNITID & "' AND FILENAME ='" & txtFileName.Text.Trim.ToString() & "'"))
            If Not IsNothing(FileName) Then
                If FileName = txtFileName.Text.Trim.ToString() Then
                    MsgBox("Selected File Already Uploaded!!")
                    Return False
                    Exit Function
                End If
            End If

            SubGetRoundoffConfig()
            Dim fd As OpenFileDialog = New OpenFileDialog()
            STRSQL = "delete from TMP_MarutisupplementaryData where Unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
            mP_Connection.Execute(STRSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            mP_Connection.BeginTrans()
            Using sr As StreamReader = New StreamReader(txtFileName.Text)

                ' Read and display lines from the file until the end of  
                ' the file is reached. q
                While (Not sr.EndOfStream)
                    line = sr.ReadLine()
                    strOriLine = line
                    batch_no = line.Replace("BATCH NO :", "§§")
                    i = batch_no.IndexOf("§§")
                    If (i > 0) Then
                        batch = batch_no.Substring(i + 2, 11).Trim()
                    End If

                    sequence_no = line.Replace("SERIAL NO :", "§§§")
                    i = sequence_no.IndexOf("§§§")
                    If (i > 0) Then
                        'serial_no = Convert.ToInt32(sequence_no.Substring(i + 3, 4).Trim())
                        serial_no = Convert.ToInt32(sequence_no.Substring(i + 3).Trim())
                    End If

                    strpartcode = line.Replace("Part No. :", "§")


                    i = strpartcode.IndexOf("§")

                    If (i > 0) Then
                        strpartcode1 = strpartcode.Substring(i + 1, 40)
                        strPartNo = strpartcode.Substring(i + 1, strpartcode1.IndexOf(" "))
                        stritemcode = Find_Value("SELECT TOP 1 item_code FROM custitem_mst WHERE UNIT_CODE='" & gstrUNITID & "' AND cust_drgno='" & strPartNo & "' and account_code='" & txtCustomerCode.Text & "'") 'and active=1
                        strpartdesc = Find_Value("SELECT TOP 1 Drg_Desc FROM custitem_mst WHERE UNIT_CODE='" & gstrUNITID & "' AND cust_drgno='" & strPartNo & "' and account_code='" & txtCustomerCode.Text & "'") 'and active=1
                        strHSNSACCODE = Find_Value("SELECT TOP 1 HSN_SAC_CODE FROM item_mst WHERE UNIT_CODE='" & gstrUNITID & "' AND item_code ='" & stritemcode & "'")
                    End If

                    stcgsttaxtype = line.Replace("CGST :", "§§§")
                    i = stcgsttaxtype.IndexOf("§§§")
                    If (i > 0) Then
                        stcgsttaxtype = stcgsttaxtype.Substring(i + 3, 2)
                        stcgsttaxper = stcgsttaxtype
                        strcgsttaxtype = "CGST" + stcgsttaxper
                    End If

                    stsgsttaxtype = line.Replace("SGST :", "§§§")
                    i = stsgsttaxtype.IndexOf("§§§")
                    If (i > 0) Then
                        stsgsttaxtype = stsgsttaxtype.Substring(i + 3, 2)
                        stsgsttaxper = stsgsttaxtype
                        strsgsttaxtype = "SGST" + stsgsttaxper
                    End If

                    If (IsNumeric(stsgsttaxper) = False Or stsgsttaxper = "0") Then
                        If stigsttaxper = 0 Then
                            stigsttaxtype = line.Replace("IGST :", "§§§")
                            i = stigsttaxtype.IndexOf("§§§")
                            If (i > 0) Then
                                stigsttaxtype = stigsttaxtype.Substring(i + 3, 2)
                                stigsttaxper = stigsttaxtype
                                strigsttaxtype = "IGST" + stigsttaxper
                            End If
                        Else
                            If stigsttaxtype = "" Then
                                stigsttaxtype = "IGST0"
                                stigsttaxper = "0.00"
                            End If
                        End If
                    Else
                        stigsttaxtype = "IGST0"
                        stigsttaxper = "0.00"
                    End If

                    If (strVal.Length = 1) Then
                        line = line.Replace(" ", ",").Replace(",,,,,,,,,,", ",").Replace(",,,,,,,,,", ",").Replace(",,,,,,,,", ",").Replace(",,,,,,,", ",").Replace(",,,,,,", ",").Replace(",,,,,", ",").Replace(",,,,", ",").Replace(",,,", ",").Replace(",,", ",")


                        intRows = UBound(Split(line, ","))
                        'strVal = strDataLine.Split(",")

                        If intRowSeperator > 1 Then
                            strRow = Mid(line, 1, intRowSeperator - 1) 'Get 1 Row at a Time
                        End If
                        Dim intPos, counter As Integer
                        Dim strFields(17) As String
                        Dim strMstFields() As String

                        Dim strMaster As String
                        Dim inti As Integer
                        Dim strAmend As String = Mid(line, 6, 5)

                        If intRows = 12 And Mid(line, 6, 5) <> "Amend" Then
                            counter = 0
                            intRowSeperator = InStr(line, ",")
                            strMaster = CStr(line).ToString
                            intPos = 0
                            counter = 11

                            While Len(strMaster) > 0
                                ReDim Preserve strMstFields(counter)
                                For inti = 0 To counter
                                    intPos = InStr(strMaster, ",")
                                    If intPos > 0 Then
                                        strMstFields(inti) = Mid(strMaster, 1, intPos - 1)
                                        strMaster = Mid(strMaster, intPos + 1)
                                    Else
                                        strMstFields(inti) = strMaster
                                        strMaster = ""
                                        Exit While
                                    End If
                                Next
                            End While

                            If IsDate(strMstFields(2)) = True Then
                                strBillNo = strMstFields(1).ToString
                                str57f2 = strMstFields(1).ToString
                                strPricDt = strMstFields(2).ToString
                                strshp = strMstFields(6).ToString
                                stracp = strMstFields(6).ToString
                                dbloldrate = Convert.ToDecimal(strMstFields(8))
                                strBilldate = strMstFields(9).ToString()
                                ' added by priti on 05.06.2019 to fix credit debit not equal
                                If blnISBasicRoundOff = True Then
                                    strbasicamt = Convert.ToDecimal(strMstFields(10))
                                Else
                                    strbasicamt = System.Math.Round(Convert.ToDecimal(strMstFields(10)), intBasicRoundOffDecimal)
                                End If
                                ' code ends here
                                If Trim(stcgsttaxper) = "%" Then
                                    stcgsttaxper = "0"
                                End If
                                If Trim(stsgsttaxper) = "%" Then
                                    stsgsttaxper = "0"
                                End If
                                ' added by priti on 05.06.2019 to fix credit debit not equal
                                If blnGSTTAXroundoff = True Then
                                    strcgstamt = strbasicamt * stcgsttaxper / 100
                                    strsgstamt = strbasicamt * stsgsttaxper / 100
                                    strigstamt = strbasicamt * stigsttaxper / 100
                                Else
                                    strcgstamt = System.Math.Round(strbasicamt * stcgsttaxper / 100, intGSTTAXroundoff_decimal)
                                    strsgstamt = System.Math.Round(strbasicamt * stsgsttaxper / 100, intGSTTAXroundoff_decimal)
                                    strigstamt = System.Math.Round(strbasicamt * stigsttaxper / 100, intGSTTAXroundoff_decimal)
                                End If
                                ' code ends here

                                dblNewRate = Convert.ToDecimal(strMstFields(9))
                                strsrvdino = strMstFields(3).ToString()
                                dbldiffrate = dblNewRate - dbloldrate
                                dbltotalamt = strbasicamt + strcgstamt + strsgstamt + strigstamt
                                If dbldiffrate > 0 Then
                                    STRSQL = "Insert Into TMP_MarutisupplementaryData (strPartNo,part_desc,item_code,strBillNo,strPricDt,strshp,stracp,stroldrate,strNewRate,DIFF_RATE,Unit_code,Ipaddress,invoiceNo,batch_no,serial_no,strBilldate,srvdino,Account_Code,basic_amount,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,hsnsaccode,TOTAL_AMOUNT ) Values ('" & strPartNo & "','" & strpartdesc & "','" & stritemcode & "','" & strBillNo & "','" & strPricDt & "','" & strshp & "','" & stracp & "','" & dbloldrate & "','" & dblNewRate & "','," & dbldiffrate & "','" & gstrUNITID & "','" & gstrIpaddressWinSck & "','" & str57f2 & "','" & batch & "'," & batch & ",'" & strBilldate & "','" & strsrvdino & "','" & txtCustomerCode.Text & "'," & strbasicamt & ",'" & strcgsttaxtype & "','" & stcgsttaxper & "'," & strcgstamt & ",'" & strsgsttaxtype & "','" & stsgsttaxper & "'," & strsgstamt & ",'" & strigsttaxtype & "','" & stigsttaxper & "'," & strigstamt & ",'" & strHSNSACCODE & "'," & dbltotalamt & " )"
                                    mP_Connection.Execute(STRSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If

                            STRSQL = ""
                        End If
                    End If
                End While
                STRSQL = ""
                mP_Connection.CommitTrans()
                'txtFileName.Text = ""

            End Using
            Return True
        Catch ex As Exception
            'RaiseException(ex)
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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

    Private Sub cmdUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpload.Click
        Try
            '' INC1557693 - Added by priti on 21 Nov 2025 for adding new format option.
            If mblnBatchFile_NewFomrat = True Then
                If chkNewFormat.Checked = True Then
                    If MsgBox("Do you want to continue with NEW format ", vbOKCancel, Me.Text) = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                Else
                    If MsgBox("Do you want to continue with OLD format ", vbOKCancel, Me.Text) = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                End If
            End If

            If chkNewFormat.Checked Then
                If MUL_SCHEDULE_UPLOAD_NEWFORMAT() = True Then
                    DISPLAY()
                Else
                    Exit Sub
                End If
            Else
                If MUL_SCHEDULE_UPLOAD() = True Then
                    DISPLAY()
                Else
                    Exit Sub
                End If
            End If

            cmdHelpDocNo.Enabled = False
            cmdCustHelp.Enabled = False
            txtCustomerCode.Enabled = False
            txtFileName.Enabled = False
            cmdFileHelp.Enabled = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub cmdClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        Try
            Call cleardata()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Function DISPLAY()
        Dim STRSQL As String = String.Empty
        Dim RSDISPLAY As ClsResultSetDB
        Dim SRNO As Integer = 0
        Dim rsTrigger As ClsResultSetDB
        Dim strReportName As String
        Dim straddress As String
        Try

            If txtDocNo.Text = "" Then
                STRSQL = "SELECT * FROM TMP_MarutisupplementaryData" &
                   " WHERE IPADDRESS = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'"

                RSDISPLAY = New ClsResultSetDB
                RSDISPLAY.GetResult(STRSQL)
                If RSDISPLAY.RowCount > 0 Then
                    RSDISPLAY.MoveFirst()
                    SSSUPPLEMENTARYDATA.MaxRows = 1

                    While Not RSDISPLAY.EOFRecord
                        With SSSUPPLEMENTARYDATA
                            SRNO = SRNO + 1
                            .SetText(ENUM_GRID.INVOICENO, .MaxRows, RSDISPLAY.GetValue("INVOICENO"))
                            .SetText(ENUM_GRID.BILLNO, .MaxRows, RSDISPLAY.GetValue("strBillNo"))
                            .SetText(ENUM_GRID.ITEMCODE, .MaxRows, RSDISPLAY.GetValue("Item_code"))
                            .SetText(ENUM_GRID.PARTNO, .MaxRows, RSDISPLAY.GetValue("strPartNo"))
                            .SetText(ENUM_GRID.DESCRIPTION, .MaxRows, RSDISPLAY.GetValue("part_desc"))
                            .SetText(ENUM_GRID.ACTQTY, .MaxRows, RSDISPLAY.GetValue("stracp"))
                            .SetText(ENUM_GRID.RATEDIFF, .MaxRows, RSDISPLAY.GetValue("DIFF_RATE"))
                            .SetText(ENUM_GRID.BASICAMT, .MaxRows, RSDISPLAY.GetValue("BASIC_AMOUNT"))
                            .SetText(ENUM_GRID.CGSTTAX, .MaxRows, RSDISPLAY.GetValue("CGSTTAX_TYPE"))
                            .SetText(ENUM_GRID.SGSTTAX, .MaxRows, RSDISPLAY.GetValue("SGSTTAX_TYPE"))
                            .SetText(ENUM_GRID.IGSTTAX, .MaxRows, RSDISPLAY.GetValue("IGSTTAX_TYPE"))
                            .SetText(ENUM_GRID.CGSTAMT, .MaxRows, RSDISPLAY.GetValue("CGST_AMT"))
                            .SetText(ENUM_GRID.SGSTAMT, .MaxRows, RSDISPLAY.GetValue("SGST_AMT"))
                            .SetText(ENUM_GRID.IGSTAMT, .MaxRows, RSDISPLAY.GetValue("IGST_AMT"))
                            .SetText(ENUM_GRID.HSNSACCODE, .MaxRows, RSDISPLAY.GetValue("HSNSACCODE"))
                            .SetText(ENUM_GRID.BATCHNO, .MaxRows, RSDISPLAY.GetValue("batch_no"))
                            .SetText(ENUM_GRID.SRVIDINO, .MaxRows, RSDISPLAY.GetValue("srvdino"))
                            .SetText(ENUM_GRID.SERIAL_NO, .MaxRows, RSDISPLAY.GetValue("SERIAL_NO"))

                            .MaxRows = .MaxRows + 1
                            RSDISPLAY.MoveNext()
                        End With
                    End While
                End If
                RSDISPLAY.ResultSetClose()
                cmdSave.Enabled = True
            Else
                STRSQL = "SELECT * FROM MarutisupplementaryData" &
                   " WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO='" & txtDocNo.Text.Trim & "'"

                RSDISPLAY = New ClsResultSetDB
                RSDISPLAY.GetResult(STRSQL)
                If RSDISPLAY.RowCount > 0 Then
                    RSDISPLAY.MoveFirst()
                    SSSUPPLEMENTARYDATA.MaxRows = 1

                    While Not RSDISPLAY.EOFRecord
                        With SSSUPPLEMENTARYDATA
                            SRNO = SRNO + 1
                            .SetText(ENUM_GRID.INVOICENO, .MaxRows, RSDISPLAY.GetValue("Ref_invNo"))
                            .SetText(ENUM_GRID.DOC_NO, .MaxRows, RSDISPLAY.GetValue("invoice_no")) : .ColHidden = False
                            .SetText(ENUM_GRID.BILLNO, .MaxRows, RSDISPLAY.GetValue("BillNo"))
                            .SetText(ENUM_GRID.ITEMCODE, .MaxRows, RSDISPLAY.GetValue("item_code"))
                            .SetText(ENUM_GRID.PARTNO, .MaxRows, RSDISPLAY.GetValue("PartNo"))
                            .SetText(ENUM_GRID.DESCRIPTION, .MaxRows, RSDISPLAY.GetValue("part_desc"))
                            .SetText(ENUM_GRID.ACTQTY, .MaxRows, RSDISPLAY.GetValue("actqty"))
                            .SetText(ENUM_GRID.RATEDIFF, .MaxRows, RSDISPLAY.GetValue("DIFF_RATE"))
                            .SetText(ENUM_GRID.BASICAMT, .MaxRows, RSDISPLAY.GetValue("BASIC_AMOUNT"))
                            .SetText(ENUM_GRID.CGSTTAX, .MaxRows, RSDISPLAY.GetValue("CGSTTAX_TYPE"))
                            .SetText(ENUM_GRID.SGSTTAX, .MaxRows, RSDISPLAY.GetValue("SGSTTAX_TYPE"))
                            .SetText(ENUM_GRID.IGSTTAX, .MaxRows, RSDISPLAY.GetValue("IGSTTAX_TYPE"))
                            .SetText(ENUM_GRID.CGSTAMT, .MaxRows, RSDISPLAY.GetValue("CGST_AMT"))
                            .SetText(ENUM_GRID.SGSTAMT, .MaxRows, RSDISPLAY.GetValue("SGST_AMT"))
                            .SetText(ENUM_GRID.IGSTAMT, .MaxRows, RSDISPLAY.GetValue("IGST_AMT"))
                            .SetText(ENUM_GRID.HSNSACCODE, .MaxRows, RSDISPLAY.GetValue("HSNSACCODE"))
                            .SetText(ENUM_GRID.BATCHNO, .MaxRows, RSDISPLAY.GetValue("batch_no"))
                            .SetText(ENUM_GRID.SRVIDINO, .MaxRows, RSDISPLAY.GetValue("srvdino"))
                            .SetText(ENUM_GRID.SERIAL_NO, .MaxRows, RSDISPLAY.GetValue("SERIAL_NO"))
                            .MaxRows = .MaxRows + 1
                            RSDISPLAY.MoveNext()
                        End With
                    End While
                End If
                RSDISPLAY.ResultSetClose()
                cmdSave.Enabled = False
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim strsql As String = ""
        Dim strchallanno As String = ""
        Dim strSlqQuery As String
        Dim strsqlHdr As String
        Dim strSqlDtl As String
        Dim strdocno As String
        Dim strupdatedocno As String
        Dim strupdateSlqQuery As String

        Try

            If SSSUPPLEMENTARYDATA.MaxRows <= 0 Then
                MsgBox(" No Row selected. ", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            SqlConnectionclass.BeginTrans()

            strsql = "SELECT ISNULL(MUL_SUPPINVOICENO ,0)+1  MUL_SUPPINVOICENO  FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "' "
            strdocno = SqlConnectionclass.ExecuteScalar(strsql)

            strsql = "Select isnull(max(Doc_No),0) as Doc_No from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and bill_flag=0 and Doc_No>" & 99000000
            strchallanno = SqlConnectionclass.ExecuteScalar(strsql)
            If strchallanno = 0 Then
                strchallanno = 99000000
            End If
            'strSlqQuery = "INSERT INTO MarutisupplementaryData  (invoice_no,PartNo,item_code,hsnsaccode,BillNo,PricDt,shipqty,actqty,OldRate,NewRate,DIFF_RATE,Unit_code,Ref_invNo,Account_Code,batch_no,serial_no,srvdino,BASIC_AMOUNT,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,PART_DESC,DOC_NO ,INVOICE_DATE,TOTAL_AMOUNT ) "
            'strSlqQuery += " SELECT  " + strchallanno + "+ ROW_NUMBER() OVER (ORDER BY line_no ) ,strPartNo,ITEM_CODE,hsnsaccode,strBillNo,strPricDt,strshp,stracp,strOldRate,strNewRate,DIFF_RATE,Unit_code,invoiceno,account_code ,batch_no ,serial_no ,srvdino,BASIC_AMOUNT,CGSTTAX_TYPE,"
            'strSlqQuery += " CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,"
            'strSlqQuery += " PART_DESC,'" & strdocno & "','" & getDateForDB(GetServerDate()) & "',total_amount from TMP_MarutisupplementaryData  where unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
            'SqlConnectionclass.ExecuteNonQuery(strSlqQuery)

            If blnSUPP_INVOICE_ONEITEM_MULTIPLE = True Then

                strSlqQuery = " INSERT INTO MarutisupplementaryData_batchwise (invoice_no,PartNo,item_code,hsnsaccode,shipqty,actqty,OldRate,NewRate,DIFF_RATE,Unit_code,Account_Code,batch_no,serial_no,BASIC_AMOUNT,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,PART_DESC,DOC_NO ,INVOICE_DATE,TOTAL_AMOUNT ) "
                strSlqQuery += " SELECT  " + strchallanno + "+ DENSE_RANK() OVER ( ORDER BY SERIAL_NO,BATCH_NO  ),strPartNo,ITEM_CODE,hsnsaccode,SUM(strshp),SUM(stracp),strOldRate,strNewRate,DIFF_RATE,Unit_code,account_code ,batch_no ,serial_no ,sum(BASIC_AMOUNT),CGSTTAX_TYPE,"
                strSlqQuery += " CGSTTAX_PER,sum(CGST_AMT),SGSTTAX_TYPE,SGSTTAX_PER,sum(SGST_AMT),IGSTTAX_TYPE,IGSTTAX_PER,sum(IGST_AMT),"
                strSlqQuery += " PART_DESC,'" & strdocno & "','" & getDateForDB(GetServerDate()) & "',sum(total_amount )from TMP_MarutisupplementaryData  where unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
                strSlqQuery += "GROUP BY strPartNo,ITEM_CODE,hsnsaccode,strOldRate,strNewRate,DIFF_RATE,Unit_code,account_code ,batch_no ,serial_no ,CGSTTAX_TYPE, "
                strSlqQuery += " CGSTTAX_PER,SGSTTAX_TYPE,SGSTTAX_PER,IGSTTAX_TYPE,IGSTTAX_PER,PART_DESC"
                SqlConnectionclass.ExecuteNonQuery(strSlqQuery)

                strSlqQuery = "INSERT INTO MarutisupplementaryData (invoice_no ,PartNo,item_code,hsnsaccode,BillNo,PricDt,shipqty,actqty,OldRate,NewRate,DIFF_RATE,Unit_code,Ref_invNo,Account_Code,batch_no,serial_no,srvdino,BASIC_AMOUNT,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,PART_DESC,DOC_NO ,INVOICE_DATE,TOTAL_AMOUNT ) "
                strSlqQuery += " SELECT  0,strPartNo,ITEM_CODE,hsnsaccode,strBillNo,strPricDt,strshp,stracp,strOldRate,strNewRate,DIFF_RATE,Unit_code,invoiceno,account_code ,batch_no ,serial_no ,srvdino,BASIC_AMOUNT,CGSTTAX_TYPE,"
                strSlqQuery += " CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,"
                strSlqQuery += " PART_DESC,'" & strdocno & "','" & getDateForDB(GetServerDate()) & "',total_amount from TMP_MarutisupplementaryData  where unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSlqQuery)

                strsqlHdr = "Insert into SupplementaryInv_hdr ("
                strsqlHdr = strsqlHdr & "DRCR,Unit_code,Location_Code,Account_Code,Cust_name,Doc_No,cust_ref,"
                strsqlHdr = strsqlHdr & "Invoice_DateFrom, Invoice_DateTo, Invoice_Date, Bill_Flag, Cancel_flag,"
                strsqlHdr = strsqlHdr & "Currency_Code,"
                strsqlHdr = strsqlHdr & "Ent_dt,Ent_UserId,Upd_dt,Upd_Userid)"
                strsqlHdr = strsqlHdr & " select distinct 'DR'," & "'" & gstrUNITID & "','" & gstrUNITID & "',Account_Code,'" & Trim(LblCustomerName.Text) & "',"
                strsqlHdr = strsqlHdr & " invoice_no,invoice_no,invoice_date,invoice_date,invoice_date,0,0,'INR',"
                strsqlHdr = strsqlHdr & " '" & getDateForDB(GetServerDate()) & "','"
                strsqlHdr = strsqlHdr & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "'"
                strsqlHdr = strsqlHdr & " from MarutisupplementaryData_batchwise (nolock) where unit_code='" & gstrUNITID & "' and doc_no=" & strdocno
                SqlConnectionclass.ExecuteNonQuery(strsqlHdr)

                strSqlDtl = "insert into supplementaryinv_dtl(Location_Code,Doc_No,LastSupplementary,SuppInvDate,Item_code,"
                strSqlDtl = strSqlDtl & "Cust_Item_Code,PrevRate,Rate,Rate_diff,Quantity,"
                strSqlDtl = strSqlDtl & " Basic_AmountDiff,Accessible_amountDiff,"
                strSqlDtl = strSqlDtl & "total_amountDiff,Ent_dt,Ent_UserID,Upd_Dt,Upd_UserID,UNIT_CODE,"
                strSqlDtl = strSqlDtl & "DIFF_CGST_AMT, DIFF_SGST_AMT,DIFF_IGST_AMT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,"
                strSqlDtl = strSqlDtl & "SGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT)"
                strSqlDtl = strSqlDtl & "select '" & gstrUNITID & "',invoice_no, 0,'" & getDateForDB(GetServerDate()) & "',item_Code,"
                strSqlDtl = strSqlDtl & "PartNo,oldrate,newrate,diff_rate,actqty,basic_amount,basic_amount,total_amount"
                strSqlDtl = strSqlDtl & ",'" & getDateForDB(GetServerDate()) & " ','"
                strSqlDtl = strSqlDtl & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & gstrUNITID & "',"
                strSqlDtl = strSqlDtl & "CGST_AMT ,SGST_AMT,IGST_AMT,CGSTTAX_TYPE,CGSTTAX_PER,SGSTTAX_TYPE,SGSTTAX_PER,IGSTTAX_TYPE,IGSTTAX_PER "
                strSqlDtl = strSqlDtl & "from MarutisupplementaryData_batchwise (nolock) where unit_code='" & gstrUNITID & "'and doc_no=" & strdocno
                SqlConnectionclass.ExecuteNonQuery(strSqlDtl)


                strupdateSlqQuery = "UPDATE M SET INVOICE_NO=MB.INVOICE_NO FROM MARUTISUPPLEMENTARYDATA M INNER JOIN MARUTISUPPLEMENTARYDATA_BATCHWISE MB ON"
                strupdateSlqQuery += " M.UNIT_CODE=MB.UNIT_CODE AND M.BATCH_NO=MB.BATCH_NO AND "
                strupdateSlqQuery += " M.SERIAL_NO=MB.SERIAL_NO AND M.ITEM_CODE=MB.ITEM_CODE AND M.PARTNO=MB.PARTNO"
                strupdateSlqQuery += " AND M.OldRate=MB.OLDRATE AND M.NEWRATE=MB.NEWRATE AND M.DIFF_RATE=MB.DIFF_RATE"
                strupdateSlqQuery += " where M.unit_code='" & gstrUNITID & "' and M.doc_no='" & strdocno & "'"
                SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)

                strupdateSlqQuery = "UPDATE SD SET DIFF_IGST_AMT=ROUND(BASIC_AMOUNTDIFF*IGST_PERCENT/100,2) , "
                strupdateSlqQuery += " DIFF_CGST_AMT = ROUND(BASIC_AMOUNTDIFF * CGST_PERCENT/100,2),"
                strupdateSlqQuery += " DIFF_SGST_AMT = ROUND(BASIC_AMOUNTDIFF * SGST_PERCENT/100,2),"
                strupdateSlqQuery += " total_amountdiff = BASIC_AMOUNTDIFF+ ROUND(BASIC_AMOUNTDIFF*IGST_PERCENT/100,2)+ROUND(BASIC_AMOUNTDIFF * CGST_PERCENT/100,2)+ROUND(BASIC_AMOUNTDIFF * SGST_PERCENT/100,2)"
                strupdateSlqQuery += " FROM SUPPLEMENTARYINV_DTL SD INNER JOIN MARUTISUPPLEMENTARYDATA_BATCHWISE M ON "
                strupdateSlqQuery += " M.UNIT_CODE=SD.UNIT_CODE AND M.INVOICE_NO=SD.doc_no"
                strupdateSlqQuery += " where M.unit_code='" & gstrUNITID & "' and M.doc_no='" & strdocno & "'"
                SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)

                strupdateSlqQuery = "UPDATE SH SET SH.TOTAL_AMOUNT=XYZ.TOTALAMOUNT , BASIC_AMOUNT =XYZ.BASIC_AMOUNTDIFF FROM "
                strupdateSlqQuery += "(SELECT M.DOC_NO,M.INVOICE_NO,M.UNIT_CODE,BATCH_NO,SERIAL_NO,SUM(M.TOTAL_AMOUNT) TOTALAMOUNT, SUM(M.BASIC_AMOUNT) BASIC_AMOUNTDIFF "
                strupdateSlqQuery += " FROM MARUTISUPPLEMENTARYDATA_BATCHWISE M (NOLOCK) INNER JOIN SUPPLEMENTARYINV_HDR SH (NOLOCK)"
                strupdateSlqQuery += " ON M.UNIT_CODE=SH.UNIT_CODE AND M.INVOICE_NO=SH.DOC_NO "
                strupdateSlqQuery += " group by batch_no,serial_no,m.unit_code,m.invoice_no ,M.DOC_NO)XYZ "
                strupdateSlqQuery += " inner join supplementaryinv_hdr sh "
                strupdateSlqQuery += " on sh.unit_code = xyz.unit_code and sh.doc_no=xyz.invoice_no "
                strupdateSlqQuery += " where XYZ.unit_code='" & gstrUNITID & "' and XYZ.doc_no='" & strdocno & "'"
                SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)

                strupdateSlqQuery = "UPDATE sh SET SH.TOTAL_AMOUNT=XYZ.TOTAL_AMOUNTDIFF "
                strupdateSlqQuery += " from (SELECT UNIT_CODE,DOC_NO,SUM(SD.TOTAL_AMOUNTDIFF )TOTAL_AMOUNTDIFF "
                strupdateSlqQuery += " from SUPPLEMENTARYINV_DTL SD group by doc_no,UNIT_CODE )XYZ "
                strupdateSlqQuery += " INNER JOIN SUPPLEMENTARYINV_HDR SH ON SH.UNIT_CODE=XYZ.UNIT_CODE   "
                strupdateSlqQuery += " AND SH.DOC_NO=XYZ.DOC_NO INNER JOIN MARUTISUPPLEMENTARYDATA_BATCHWISE M ON "
                strupdateSlqQuery += " M.UNIT_CODE = XYZ.UNIT_CODE And XYZ.DOC_NO = M.INVOICE_NO "
                strupdateSlqQuery += " where M.unit_code='" & gstrUNITID & "' and M.doc_no='" & strdocno & "'"
                SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)

                strupdatedocno = "UPDATE SALES_PARAMETER SET MUL_SUPPINVOICENO=" & strdocno & "WHERE UNIT_CODE='" & gstrUNITID & "'"

                SqlConnectionclass.ExecuteNonQuery(strupdatedocno)

                strsql = ""
                strsql = "INSERT INTO MARUTI_SUPPINVOICE_FILE_HDR (FILENAME,UNIT_CODE,DOCNO,ENT_DT) Values ('" & txtFileName.Text.Trim.ToString() & "','" & gstrUNITID & "','" & strdocno & "',getdate())"
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Else

                strSlqQuery = "INSERT INTO MarutisupplementaryData (invoice_no,PartNo,item_code,hsnsaccode,BillNo,PricDt,shipqty,actqty,OldRate,NewRate,DIFF_RATE,Unit_code,Ref_invNo,Account_Code,batch_no,serial_no,srvdino,BASIC_AMOUNT,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,PART_DESC,DOC_NO ,INVOICE_DATE,TOTAL_AMOUNT ) "
                strSlqQuery += " SELECT  " + strchallanno + "+ ROW_NUMBER() OVER (ORDER BY line_no ) strPartNo,ITEM_CODE,hsnsaccode,strBillNo,strPricDt,strshp,stracp,strOldRate,strNewRate,DIFF_RATE,Unit_code,invoiceno,account_code ,batch_no ,serial_no ,srvdino,BASIC_AMOUNT,CGSTTAX_TYPE,"
                strSlqQuery += " CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,"
                strSlqQuery += " PART_DESC,'" & strdocno & "','" & getDateForDB(GetServerDate()) & "',total_amount from TMP_MarutisupplementaryData  where unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSlqQuery)

                strsqlHdr = "Insert into SupplementaryInv_hdr ("
                strsqlHdr = strsqlHdr & "DRCR,Unit_code,Location_Code,Account_Code,Cust_name,Doc_No,cust_ref,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,Item_Code,"
                strsqlHdr = strsqlHdr & "Cust_Item_Code,Currency_Code,Rate,Basic_Amount,Accessible_amount,"
                strsqlHdr = strsqlHdr & "total_amount,Ent_dt,Ent_UserId,Upd_dt,Upd_Userid"
                strsqlHdr = strsqlHdr & ",HSN_SAC_CODE,CGSTTXRT_TYPE,CGST_PERCENT,CGST_AMT,SGSTTXRT_TYPE,SGST_PERCENT,SGST_AMT,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT)"
                strsqlHdr = strsqlHdr & " select 'DR'," & "'" & gstrUNITID & "','" & gstrUNITID & "',Account_Code,'" & Trim(LblCustomerName.Text) & "',"
                strsqlHdr = strsqlHdr & " invoice_no,invoice_no,invoice_date,invoice_date,invoice_date,0,0,item_code,partno,'INR',diff_rate,basic_amount,basic_amount,"
                strsqlHdr = strsqlHdr & " total_amount,'" & getDateForDB(GetServerDate()) & "','"
                strsqlHdr = strsqlHdr & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "',HSNSACCODE,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,"
                strsqlHdr = strsqlHdr & "SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT from MarutisupplementaryData (nolock) where unit_code='" & gstrUNITID & "' and doc_no=" & strdocno
                SqlConnectionclass.ExecuteNonQuery(strsqlHdr)

                strSqlDtl = "insert into supplementaryinv_dtl(Location_Code,Doc_No,RefDoc_No,RefDoc_Date,LastSupplementary,SuppInvDate,Item_code,"
                strSqlDtl = strSqlDtl & "Cust_Item_Code,PrevRate,Rate,Rate_diff,Quantity,"
                strSqlDtl = strSqlDtl & " Basic_AmountDiff,Accessible_amountDiff,"
                strSqlDtl = strSqlDtl & "total_amountDiff,Ent_dt,Ent_UserID,Upd_Dt,Upd_UserID,UNIT_CODE,"
                strSqlDtl = strSqlDtl & "DIFF_CGST_AMT, DIFF_SGST_AMT,DIFF_IGST_AMT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,"
                strSqlDtl = strSqlDtl & "SGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT)"
                strSqlDtl = strSqlDtl & "select '" & gstrUNITID & "',invoice_no, Ref_invNo,pricdt,0,'" & getDateForDB(GetServerDate()) & "',item_Code,"
                strSqlDtl = strSqlDtl & "PartNo,oldrate,newrate,diff_rate,actqty,basic_amount,basic_amount,total_amount"
                strSqlDtl = strSqlDtl & ",'" & getDateForDB(GetServerDate()) & " ','"
                strSqlDtl = strSqlDtl & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & gstrUNITID & "',"
                strSqlDtl = strSqlDtl & "CGST_AMT ,SGST_AMT,IGST_AMT,CGSTTAX_TYPE,CGSTTAX_PER,SGSTTAX_TYPE,SGSTTAX_PER,IGSTTAX_TYPE,IGSTTAX_PER "
                strSqlDtl = strSqlDtl & "from MarutisupplementaryData (nolock) where unit_code='" & gstrUNITID & "'and doc_no=" & strdocno
                SqlConnectionclass.ExecuteNonQuery(strSqlDtl)

                strupdatedocno = "update sales_parameter set MUL_SUPPINVOICENO=" & strdocno & "where unit_code='" & gstrUNITID & "'"
                SqlConnectionclass.ExecuteNonQuery(strupdatedocno)

                strsql = ""
                strsql = "INSERT INTO MARUTI_SUPPINVOICE_FILE_HDR (FILENAME,UNIT_CODE,DOCNO,ENT_DT) Values ('" & txtFileName.Text.Trim.ToString() & "','" & gstrUNITID & "','" & strdocno & "',getdate())"
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If


            SqlConnectionclass.CommitTran()
            MsgBox("Invoices locked successfully with Transaction No : " + strdocno, MsgBoxStyle.Information, ResolveResString(100))
            Call cleardata()

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub cleardata()
        Try

            txtCustomerCode.Text = ""
            txtDocNo.Text = ""
            txtFileName.Text = ""
            SSSUPPLEMENTARYDATA.MaxRows = 0
            cmdUpload.Enabled = True
            cmdHelpDocNo.Enabled = True
            cmdCustHelp.Enabled = True
            txtCustomerCode.Enabled = True
            txtFileName.Enabled = True
            cmdFileHelp.Enabled = True

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub frmMKTTRN0101_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

    End Sub
End Class