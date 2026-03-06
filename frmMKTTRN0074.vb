Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports System.Data
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmMKTTRN0074
#Region "Comments"
    '***************************************************************************************
    'Copyright       : MIND  Ltd.
    'Module          : frmMKTTRN0074 - Intra Material Transfer Note
    'Author          : Parveen Kumar
    'Creation Date   : 04 Jul 2012
    'Description     : Intra Material Transfer Note
    '***************************************************************************************
	'Revised By     :   Shubhra Verma
    'Revised On     :   26 Sep 2012
    'Issue ID       :   10281382 
    'Description    :   Grid was showing Inactive data as well.
    '***************************************************************************************
    'Revised By     :   Shubhra Verma
    'Revised On     :   01 Oct 2012
    'Issue ID       :   10283347 
    'Description    :   Remove active condition from View mode
    '***************************************************************************************
#End Region
#Region "Form level variable Declarations"
    Private GstrSqlConnectionString As String
    Dim MINTFORMINDEX As Integer
    Dim Conn As SqlConnection
    Dim SqlCmd As SqlCommand
    Dim SqlDR As SqlDataReader
    Dim SqlAdp As SqlDataAdapter
    Dim mStrSql As String = String.Empty
    Public mstrItemCode As String
    Dim intMaxLoop As Integer = 0
    Dim intLoopCounter As Integer = 0
    Dim mStrCurrency As String = String.Empty
    Dim intPerValue As Integer = 0
    Dim intRowCount As Integer = 0
    Dim VAR_ITEMCODE As Object = Nothing
    Dim VAR_CUSTPARTNO As Object = Nothing
    Dim VAR_RATE As Object = Nothing
    Dim VAR_QUANTITY As Object = Nothing
    Dim VAR_CUSTPARTNODESC As Object = Nothing
    Dim VAR_CURRSTOCK As Object = Nothing
    Dim VAR_LOCATION_CODE As Object = Nothing
    Dim VAR_ITEMGRPCODE As Object = Nothing
    Dim VAR_ITEMGLCODE As Object = Nothing
    Dim VAR_ITEMSLCODE As Object = Nothing
    Dim mStrGLCode As String = String.Empty
    Dim mStrSLCode As String = String.Empty
    Dim mStrCreditTerms As String = String.Empty
    Dim mSchTypeArr() As String
    Dim mStrMSG As String = String.Empty
    Dim strYYYYmm As String = String.Empty
    Private Structure PrevScheduleQty
        Dim ItemCode As String
        Dim CustItemCode As String
        Dim PrevQuantity As Decimal
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public ScheduleType() As Char
    End Structure
    Private Structure CurScheduleQty
        Dim ItemCode As String
        Dim CustItemCode As String
        Dim CurQuantity As Decimal
    End Structure
    Private marrPrevQty() As PrevScheduleQty
    Private marrCurQty() As CurScheduleQty
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
    Private Sub frmMKTTRN0074_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = MINTFORMINDEX
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub frmMKTTRN0074_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            GstrSqlConnectionString = "SERVER=" & gstrCONNECTIONSERVER & ";DATABASE=" & gstrDatabaseName & ";UID=" & gstrCONNECTIONUSER$ & ";PASSWORD=" & gstrCONNECTIONPASSWORD$ & ""
            Conn = New SqlConnection(GstrSqlConnectionString)
            'Call FillLabelFromResFile(Me)
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
            CmdCustomerCode.Image = My.Resources.ico111.ToBitmap
            cmdSoNo.Image = My.Resources.ico111.ToBitmap
            cmdHlpToUnit.Image = My.Resources.ico111.ToBitmap
            lblUnitCode.Text = gstrUNITID
            lblUnitDesc.Text = gstrUNITDESC
            CmdButtonGrp.ShowButtons(True, False, False, True)
            InitializeSpreed()
            AddBlankRow()
            RefreshForm()
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
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_QUANTITY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
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
            If CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                txtIMNTNo.Text = "" : txtIMNTNo.Enabled = True : txtIMNTNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdIMNTNo.Enabled = True
                DTPIMNT.Enabled = False : DTPIMNT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtCustomerCode.Text = "" : txtCustomerCode.Enabled = False : txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblCustomerDesc.Text = ""
                txtSoNo.Text = "" : txtSoNo.Enabled = False : txtSoNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblAmendementNo.Text = ""
                lblARDrNote.Text = ""
                CmdCustomerCode.Enabled = False
                cmdSoNo.Enabled = False
                CmdItems.Enabled = False
                mStrGLCode = ""
                mStrSLCode = ""
                SprIMNTGrid.Enabled = False : SprIMNTGrid.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                SprIMNTGrid.MaxRows = 0
                cmdHlpToUnit.Enabled = False
                lblReceivingUnitCode.Text = ""
                lblReceivingUnitName.Text = ""
            ElseIf CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                txtIMNTNo.Text = "" : txtIMNTNo.Enabled = False : txtIMNTNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdIMNTNo.Enabled = False
                DTPIMNT.Enabled = False : DTPIMNT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtCustomerCode.Text = "" : txtCustomerCode.Enabled = True : txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblCustomerDesc.Text = ""
                txtSoNo.Text = "" : txtSoNo.Enabled = True : txtSoNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblAmendementNo.Text = ""
                lblARDrNote.Text = ""
                CmdCustomerCode.Enabled = True
                cmdSoNo.Enabled = True
                CmdItems.Enabled = True
                mStrGLCode = ""
                mStrSLCode = ""
                SprIMNTGrid.Enabled = True : SprIMNTGrid.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                SprIMNTGrid.MaxRows = 0
                cmdHlpToUnit.Enabled = True
                lblReceivingUnitCode.Text = ""
                lblReceivingUnitName.Text = ""
            ElseIf CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            End If
            CmdButtonGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub ShowHelp(ByVal pstrQuery As String, ByRef pctlCode As System.Windows.Forms.TextBox, Optional ByRef pctlDate As System.Windows.Forms.DateTimePicker = Nothing, Optional ByRef pctlDesc As System.Windows.Forms.Label = Nothing)
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
                    pctlCode.Text = StrHelp(0)
                    If pctlCode.Name.ToString.Trim.ToUpper = "TXTIMNTNO" Then
                        lblReceivingUnitCode.Text = StrHelp(2)
                        lblReceivingUnitName.Text = FetchUnitName(StrHelp(2))
                    End If
                    If pctlCode.Name.ToString.Trim.ToUpper = "TXTCUSTOMERCODE" Then
                        mStrGLCode = StrHelp(2).ToString.Trim
                        mStrSLCode = StrHelp(3).ToString.Trim
                    End If
                    If Not pctlDate Is Nothing Then
                        If pctlDate.Name = "DTPIMNT" Then
                            pctlDate.Value = StrHelp(1)
                        End If
                    End If
                    If Not pctlCode.Name.ToString.Trim.ToUpper = "TXTIMNTNO" Then
                        If pctlDesc.Name.ToString.Trim.ToUpper = "LBLCUSTOMERDESC" Then
                            pctlDesc.Text = StrHelp(1)
                        End If
                        If pctlDesc.Name.ToString.Trim.ToUpper = "LBLAMENDEMENTNO" Then
                            pctlDesc.Text = StrHelp(1)
                        End If
                        
                    End If
                Else
                    MessageBox.Show("Record not found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If

            End If
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MessageBox.Show(EX.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub GetPaymentdates()
        Dim objCreditTerm As New prj_CreditTerm.clsCR_Term_Resolver
        Try
            Dim ProxDate As Object = Nothing
            Dim BaseDate As Object = Nothing
            Dim CrTermLineNo As String = String.Empty
            mStrSql = "select a.CrTrm_TermId,b.CrTrmD_lineNo,b.CrTrmD_baseDate,b.CrTrmD_fixDate,b.CrTrmD_minDueDays, " & _
                    " b.CrTrmD_proxPeriod,b.CrTrmD_proxDay,b.CrTrmD_percVal,b.CrTrmD_fixedVal " & _
                    " from Gen_CreditTrmMaster a,Gen_CreditTrmDtl b " & _
                    " where a.unit_code = '" & gstrUNITID.ToString.Trim & "'  " & _
                    " and a.CrTrm_TermId='" & mStrCreditTerms.ToString.Trim & "'   " & _
                    " and a.CrTrm_TermId=b.CrTrmD_TermId " & _
                    " and a.unit_code = b.unit_code "
            SqlCmd = New SqlCommand
            With SqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = mStrSql
                SqlDR = SqlCmd.ExecuteReader()
                If SqlDR.Read = True Then
                    CrTermLineNo = SqlDR("CrTrmD_lineNo")
                    DtpBasicDueDate.Value = DateAdd(Microsoft.VisualBasic.DateInterval.Day, SqlDR("CrTrmD_minDueDays"), DTPIMNT.Value)
                    ProxDate = objCreditTerm.getProxDate(gstrUNITID.ToString.Trim, DtpBasicDueDate.Value, mStrCreditTerms.ToString.Trim, CrTermLineNo, mP_Connection)
                    DtpBasicPayDate.Value = setDateFormat(ProxDate)
                Else
                    DtpBasicDueDate.Format = DateTimePickerFormat.Custom
                    DtpBasicDueDate.CustomFormat = gstrDateFormat
                    DtpBasicDueDate.Value = GetServerDate()
                    DtpBasicDueDate.Enabled = False
                    DtpBasicPayDate.Format = DateTimePickerFormat.Custom
                    DtpBasicPayDate.CustomFormat = gstrDateFormat
                    DtpBasicPayDate.Value = GetServerDate()
                    DtpBasicPayDate.Enabled = False
                End If
                If SqlDR.IsClosed = False Then SqlDR.Close()
                SqlCmd.Connection.Close()
                SqlCmd.Dispose()
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            If SqlDR.IsClosed = False Then SqlDR.Close()
            SqlCmd.Connection.Close()
            SqlCmd.Dispose()
        Finally
            objCreditTerm = Nothing
        End Try
    End Sub
    Private Function FetchUnitName(ByVal pstrUnitID As String) As String
        Try
            mStrSql = "SELECT UNT_UNITNAME  FROM GEN_UNITMASTER WHERE Unt_CodeID='" + pstrUnitID + "' "
            Return Convert.ToString(SqlConnectionclass.ExecuteScalar(SqlConnectionclass.GetConnection, CommandType.Text, mStrSql))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, ResolveResString(100))
        End Try

    End Function
#End Region
#Region "Form Control Events"
    Private Sub CmdButtonGrp_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdButtonGrp.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    RefreshForm()
                    txtCustomerCode.Focus()

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If Validate_data() = True Then
                        If SaveData() = False Then
                            MessageBox.Show("Due to some error record could not saved please contact to administrator.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Else
                            MessageBox.Show(mStrMSG, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            CmdButtonGrp.Revert()
                            RefreshForm()
                        End If
                    Else
                        MessageBox.Show(mStrMSG, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    CmdButtonGrp.Revert()
                    RefreshForm()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                    If Fetchdata_rpt() = True Then
                        Dim rdCDS As ReportDocument
                        Dim strRepPath As String = String.Empty
                        Dim RepViewer As New eMProCrystalReportViewer
                        Dim reptitle As String = String.Empty
                        Dim strReportName As String = String.Empty
                        Dim strRptTitle As String = String.Empty

                        rdCDS = RepViewer.GetReportDocument
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
                        strReportName = "\Reports\IMTNRpt_" & strReportName & ".rpt"
                        If Not CheckFile(strReportName) Then
                            strReportName = "\Reports\IMTNRpt.rpt"
                        End If
                        strRepPath = My.Application.Info.DirectoryPath & strReportName
                        rdCDS.Load(strRepPath)
                        strRptTitle = "Intra Material Transfer Note"
                        With rdCDS
                            .SetParameterValue("reptitle", strRptTitle)
                            .SetParameterValue("cname", gstrCOMPANY.ToString.Trim)
                            .SetParameterValue("unit", gstrUNITDESC.ToString.Trim)
                            .RecordSelectionFormula = "{Temp_IMTN_ReportData.IpAddress}='" & gstrIpaddressWinSck & "' AND {Temp_IMTN_ReportData.unit_code} = '" & gstrUNITID & "'"
                        End With

                        RepViewer.Show()
                    Else
                        MessageBox.Show("Due to some error report could not be generate please contact to administrator.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub CmdIMNTNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdIMNTNo.Click, CmdCustomerCode.Click, cmdSoNo.Click
        Try
            If sender.Equals(CmdIMNTNo) Then
                mStrSql = "select Doc_No,doc_date,Recv_Unit_Code AS Receiving_Unit from imtn_hdr WHERE UNIT_CODE = '" & gstrUNITID & "'"
                Call ShowHelp(mStrSql, txtIMNTNo, DTPIMNT)
                If txtIMNTNo.Text.ToString.Trim.Length > 0 Then

                    If PopulateData() = False Then
                        MessageBox.Show("Record not found", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If
                End If
            ElseIf sender.Equals(CmdCustomerCode) Then
                mStrGLCode = ""
                mStrSLCode = ""
                mStrSql = "select i.customer_code,c.Cust_Name,c.account_ledger,c.account_subledger from Customer_mst c inner join IMTN_UnitCustomer_Map i " & _
                        " on c.Customer_Code = i.Customer_Code and c.UNIT_CODE = i.Unit_Code " & _
                        " WHERE C.UNIT_CODE = '" & gstrUNITID & "'"
                Call ShowHelp(mStrSql, txtCustomerCode, , lblCustomerDesc)
            ElseIf sender.Equals(cmdSoNo) Then
                If Not txtCustomerCode.Text.ToString.Trim.Length = 0 Then
                    mStrSql = "SELECT   DISTINCT b.cust_ref,A.Amendment_No FROM     cust_ord_hdr a, cust_ord_dtl b " & _
                                " WHERE(a.unit_code = b.unit_code) AND b.account_code = '" & txtCustomerCode.Text.ToString.Trim & "' " & _
                                " AND b.active_flag = 'A' AND a.unit_code = '" & gstrUNITID.ToString.Trim & "' AND a.account_code = b.account_code " & _
                                " AND a.cust_ref = b.cust_ref AND a.amendment_no = b.amendment_no  " & _
                                " AND a.authorized_flag = 1 AND a.po_type IN ('O','S','M')  " & _
                                " AND a.valid_date >= '" & getDateForDB(DTPIMNT.Value) & "' AND effect_date <= '" & getDateForDB(DTPIMNT.Value) & "' " & _
                                " ORDER BY b.cust_ref "
                    Call ShowHelp(mStrSql, txtSoNo, , lblAmendementNo)
                Else
                    MessageBox.Show("Please select customer code first.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub CmdItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdItems.Click
        Try
            Dim varColumns As Object = Nothing
            Dim strStockLocation As String = String.Empty
            Dim strItemNotIn As String = String.Empty

            If SprIMNTGrid.MaxRows > 0 Then
                intMaxLoop = SprIMNTGrid.MaxRows
                strItemNotIn = ""
                For intLoopCounter = 1 To intMaxLoop
                    With SprIMNTGrid
                        VAR_ITEMCODE = Nothing
                        Call .GetText(ENUMDETAILS.VAL_ITEMCODE, intLoopCounter, VAR_ITEMCODE)
                        If Len(Trim(strItemNotIn)) > 0 Then
                            strItemNotIn = Trim(strItemNotIn) & ",'" & Trim(VAR_ITEMCODE) & "'"
                        Else
                            strItemNotIn = "'" & Trim(VAR_ITEMCODE) & "'"
                        End If
                    End With
                Next
            End If
            If strItemNotIn = "''" Then strItemNotIn = ""
            If strItemNotIn.ToString.Trim.Length > 0 Then
                mstrItemCode = frmMKTTRN0074A.CheckSoItems(txtCustomerCode.Text.ToString.Trim, txtSoNo.Text.ToString.Trim, lblAmendementNo.Text.ToString.Trim, strItemNotIn)
            Else
                mstrItemCode = frmMKTTRN0074A.CheckSoItems(txtCustomerCode.Text.ToString.Trim, txtSoNo.Text.ToString.Trim, lblAmendementNo.Text.ToString.Trim)
            End If
            If mstrItemCode = Nothing Then Exit Sub
            If mstrItemCode.ToString.Trim = "" Then Exit Sub
            mstrItemCode = Mid(mstrItemCode, 1, Len(mstrItemCode) - 1)
            If GetSOHdrValues() = False Then
                MessageBox.Show("Please check currency, credit terms and per value defint in selected SO.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Else
                GetPaymentdates()
                If DisplayDetailsInSpread(mStrCurrency) = False Then
                    MessageBox.Show("Please check currency, credit terms and per value defint in selected SO.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Me.txtCustomerCode.Enabled = True
                    Me.txtSoNo.Enabled = True
                    Me.CmdCustomerCode.Enabled = True
                    Me.cmdSoNo.Enabled = True
                Else
                    Me.txtCustomerCode.Enabled = False
                    Me.txtSoNo.Enabled = False
                    Me.CmdCustomerCode.Enabled = False
                    Me.cmdSoNo.Enabled = False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub txtIMNTNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtIMNTNo.KeyDown, txtCustomerCode.KeyDown, txtSoNo.KeyDown
        Try
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            If Shift <> 0 Then Exit Sub
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                If sender.Equals(txtIMNTNo) Then
                    Call CmdIMNTNo_Click(CmdIMNTNo, New System.EventArgs())
                ElseIf sender.Equals(txtCustomerCode) Then
                    Call CmdIMNTNo_Click(CmdCustomerCode, New System.EventArgs())
                ElseIf sender.Equals(txtSoNo) Then
                    Call CmdIMNTNo_Click(cmdSoNo, New System.EventArgs())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub txtIMNTNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtIMNTNo.Validating, txtCustomerCode.Validating, txtSoNo.Validating
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
            If sender.Equals(txtCustomerCode) Then
                If txtCustomerCode.Text.ToString.Trim.Length > 0 Then
                    mStrSql = "select Customer_code , Cust_Name from customer_Mst where customer_code='" & txtCustomerCode.Text & "' and   unit_code = '" & gstrUNITID.ToString.Trim & "' and  ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                    SqlCmd = New SqlCommand
                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandType = CommandType.Text
                        .CommandText = mStrSql
                        SqlDR = SqlCmd.ExecuteReader()
                        If SqlDR.Read = True Then
                            lblCustomerDesc.Text = SqlDR("cust_name").ToString.Trim
                        Else
                            MessageBox.Show("Customer code does not exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            txtCustomerCode.Text = ""
                            txtCustomerCode.Focus()
                        End If
                    End With
                    If SqlDR.IsClosed = False Then SqlDR.Close()
                    SqlCmd.Connection.Close()
                    SqlCmd.Dispose()
                End If
            End If
            If sender.Equals(txtSoNo) Then
                If txtSoNo.Text.ToString.Trim.Length > 0 Then
                    mStrSql = "SELECT   DISTINCT b.cust_ref,A.Amendment_No FROM     cust_ord_hdr a, cust_ord_dtl b " & _
                                   " WHERE(a.unit_code = b.unit_code) AND b.account_code = '" & txtCustomerCode.Text.ToString.Trim & "' " & _
                                   " AND b.active_flag = 'A' AND a.unit_code = '" & gstrUNITID.ToString.Trim & "' AND a.account_code = b.account_code " & _
                                   " AND a.cust_ref = b.cust_ref AND a.amendment_no = b.amendment_no  " & _
                                   " AND a.authorized_flag = 1 AND a.po_type IN ('O','S','M')  " & _
                                   " AND a.valid_date >= '" & getDateForDB(DTPIMNT.Value) & "' AND effect_date <= '" & getDateForDB(DTPIMNT.Value) & "' " & _
                                   " and b.cust_ref = '" & txtSoNo.Text.ToString.Trim & "' " & _
                                   " ORDER BY b.cust_ref "
                    SqlCmd = New SqlCommand
                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandType = CommandType.Text
                        .CommandText = mStrSql
                        SqlDR = SqlCmd.ExecuteReader()
                        If SqlDR.Read = True Then
                            lblAmendementNo.Text = SqlDR("Amendment_No").ToString.Trim
                        Else
                            MessageBox.Show("SO does not exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            lblAmendementNo.Text = ""
                            txtSoNo.Text = ""
                            txtSoNo.Focus()
                        End If
                    End With
                    If SqlDR.IsClosed = False Then SqlDR.Close()
                    SqlCmd.Connection.Close()
                    SqlCmd.Dispose()

                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
    Private Sub txtIMNTNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIMNTNo.KeyPress, txtCustomerCode.KeyPress, txtSoNo.KeyPress
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
            If CmdButtonGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then Exit Sub
            If txtIMNTNo.Text.ToString.Trim.Length = 0 Then
                Call RefreshForm()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub SprIMNTGrid_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles SprIMNTGrid.DblClick
        Try
            Dim msgreturnval As String = ""
            With SprIMNTGrid
                If e.row > 0 AndAlso e.col = 0 AndAlso (Me.CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then
                    msgreturnval = MessageBox.Show("Delete Row?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If msgreturnval = MsgBoxResult.Yes Then
                        .Row = e.row : .Row2 = e.row : .BlockMode = True : .Action = FPSpreadADO.ActionConstants.ActionDeleteRow : .MaxRows = .MaxRows - 1 : .BlockMode = False
                    End If
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
    Private Sub cmdHlpToUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHlpToUnit.Click
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim strUnit() As String
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strQry = " SELECT PARTNER_UNIT AS Receiving_Unit,PARTNER_UNIT_DESC  AS Unit_Name FROM INTRAMATERIALUNITMAPPING WHERE UNIT_CODE='" & gstrUNITID & "' "
        strUnit = IMNTHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
        If Not (UBound(strUnit) = -1) Then
            If (Len(strUnit(0)) >= 1) And strUnit(0) = "0" Then
                MsgBox("No Receiving Unit is defined.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Sub
            Else
                lblReceivingUnitCode.Text = strUnit(0)
                lblReceivingUnitName.Text = strUnit(1)
            End If
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
#End Region
#Region "Functions"
    Private Function DisplayDetailsInSpread(ByRef pstrCurrency As String) As Boolean
        Try

            Dim sqlAdp As SqlDataAdapter = Nothing
            Dim sqlDataSet As DataSet = Nothing
            Dim varItemAlready As Object = Nothing
            Dim inti As Integer = 0
            Dim intRecordCount As Integer = 0
            Dim intRowCount As Integer = 0

            Select Case Me.CmdButtonGrp.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    mStrSql = "Select A.Item_Code,A.Cust_DrgNo,A.Rate,A.cust_drg_desc,B.CUR_BAL,B.LOCATION_CODE,B.ITEM_GLGRP,B.ITEM_GLCODE,B.ITEM_SLCODE " & _
                            " from Cust_ord_dtl A INNER JOIN TEMP_INTRAMAT B ON A.ITEM_CODE = B.ITEM_CODE " & _
                            " AND A.UNIT_CODE = B.UNIT_CODE " & _
                            " WHERE A.UNIT_CODE='" + gstrUNITID.ToString.Trim + "' AND  " & _
                            " A.Account_Code ='" & txtCustomerCode.Text.ToString.Trim & "' " & _
                            " and A.Cust_ref ='" & txtSoNo.Text.ToString.Trim & "' " & _
                            " and A.Amendment_No = '" & lblAmendementNo.Text.ToString.Trim & "'and A.Active_flag ='A' " & _
                            " and A.Cust_DrgNo in(" & mstrItemCode & ") " & _
                            " AND B.LOCATION_CODE IN ('01M1','01P1') AND B.CUR_BAL > 0 " & _
                            "  and b.IP_Address = '" & gstrIpaddressWinSck.ToString.Trim & "'"
            End Select
            sqlAdp = New SqlDataAdapter(mStrSql, SqlConnectionclass.GetConnection)
            sqlDataSet = New DataSet
            sqlAdp.Fill(sqlDataSet)

            If sqlDataSet.Tables(0).Rows.Count > 0 Then
                If CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    With SprIMNTGrid
                        If .MaxRows > 0 Then
                            varItemAlready = Nothing
                            Call .GetText(ENUMDETAILS.VAL_ITEMCODE, 1, varItemAlready)
                            If Len(Trim(varItemAlready)) > 0 Then
                                inti = .MaxRows + 1
                                .MaxRows = .MaxRows + intRecordCount
                                intRecordCount = .MaxRows
                            Else
                                inti = 1
                            End If
                        Else
                            inti = 1
                            .MaxRows = intRecordCount
                        End If
                    End With
                Else
                    inti = 1
                End If
                intLoopCounter = inti
                For intRowCount = 0 To sqlDataSet.Tables(0).Rows.Count - 1
                    AddBlankRow()
                    With Me.SprIMNTGrid
                        Select Case Me.CmdButtonGrp.Mode
                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                                Call .SetText(ENUMDETAILS.VAL_ITEMCODE, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("Item_Code").ToString)
                                Call .SetText(ENUMDETAILS.VAL_CUSTPARTNO, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("Cust_DrgNo").ToString.Trim)
                                Call .SetText(ENUMDETAILS.VAL_RATE, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("Rate").ToString.Trim)
                                Call .SetText(ENUMDETAILS.VAL_QUANTITY, intLoopCounter, "0.00")
                                Call .SetText(ENUMDETAILS.VAL_CUSTPARTNODESC, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("cust_drg_desc").ToString.Trim)
                                Call .SetText(ENUMDETAILS.VAL_CURRSTOCK, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("CUR_BAL").ToString.Trim)
                                Call .SetText(ENUMDETAILS.VAL_LOCATION, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("LOCATION_CODE").ToString.Trim)
                                Call .SetText(ENUMDETAILS.VAL_ITEMGRPCODE, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("ITEM_GLGRP").ToString.Trim)
                                Call .SetText(ENUMDETAILS.VAL_ITEMGLCODE, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("ITEM_GLCODE").ToString.Trim)
                                Call .SetText(ENUMDETAILS.VAL_ITEMSLCODE, intLoopCounter, sqlDataSet.Tables(0).Rows(intRowCount)("ITEM_SLCODE").ToString.Trim)
                        End Select
                    End With
                    intLoopCounter = intLoopCounter + 1
                Next intRowCount
                Return True
            End If
            If SprIMNTGrid.MaxRows > 3 Then
                SprIMNTGrid.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    Private Function GetSOHdrValues() As Boolean
        Try
            If txtSoNo.Text.ToString.Trim.Length > 0 Then
                mStrSql = "Select max(Order_date),Currency_code,PerValue,term_payment from Cust_ord_hdr " & _
                " WHERE UNIT_CODE='" + gstrUNITID.ToString.Trim + "' AND  Account_Code='" & txtCustomerCode.Text.ToString.Trim & "'  " & _
                " and Cust_Ref ='" & txtSoNo.Text.ToString.Trim & "'and Amendment_No ='" & lblAmendementNo.Text.ToString.Trim & "' " & _
                " and active_flag = 'A' Group by currency_code,PerValue,term_payment"
                SqlCmd = New SqlCommand
                With SqlCmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .CommandType = CommandType.Text
                    .CommandText = mStrSql
                    SqlDR = SqlCmd.ExecuteReader()
                    If SqlDR.Read = True Then
                        mStrCurrency = SqlDR("Currency_code")
                        intPerValue = SqlDR("PerValue")
                        mStrCreditTerms = SqlDR("term_payment")
                    End If
                    If SqlDR.IsClosed = False Then SqlDR.Close()
                    SqlCmd.Connection.Close()
                    SqlCmd.Dispose()
                End With
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    Private Function GetNewDocumentNumber(ByVal DocumentArea As String, ByVal DocumentDate As Date, ByVal UnitCode As String, Optional ByRef DocumentId As String = "", Optional ByRef DocumentCategory As String = "", Optional ByRef DocumentSubCategory As String = "") As String
        Try
            Dim strResultSet As String = String.Empty
            Dim strDocumentArea As String = String.Empty, strDocumentNumber As String = String.Empty
            Dim lngDocNumLen As Integer = 0
            mStrSql = "SELECT DocG_Prefix,  isnull(DocG_RSNo,0) "
            mStrSql = mStrSql & "FROM Gen_DocGenMst "
            mStrSql = mStrSql & "WHERE DocG_DocArea = '" & DocumentArea & "' AND '" & getDateForDB(DocumentDate) & "' BETWEEN DocG_StDate AND DocG_EndDt " & "AND DocG_UntCodeId='" & UnitCode & "'"
            If DocumentId <> "" Then
                mStrSql = mStrSql & " AND DocG_DocID='" & DocumentId & "' "
            End If
            If DocumentCategory <> "" Then
                mStrSql = mStrSql & " AND DocG_DocCat='" & DocumentCategory & "' "
            End If
            If DocumentSubCategory <> "" Then
                mStrSql = mStrSql & " AND DocG_DocSubCat='" & DocumentSubCategory & "'"
            End If
            mStrSql = "Set dateformat 'dmy' " & vbCrLf & mStrSql
            SqlCmd = New SqlCommand(mStrSql, SqlConnectionclass.GetConnection)
            SqlDR = SqlCmd.ExecuteReader()
            If SqlDR.Read Then
                strDocumentNumber = (SqlDR(1) + 1).ToString.Trim
                lngDocNumLen = Len(strDocumentNumber)
                If lngDocNumLen < 5 Then
                    strDocumentNumber = New String("0", 5 - (lngDocNumLen)) & strDocumentNumber
                End If
                strResultSet = SqlDR(0).ToString.Trim & strDocumentNumber
                GetNewDocumentNumber = "Y¤" & strResultSet
            Else
                GetNewDocumentNumber = "N¤»Document Number Generation FAILED"
            End If
            SqlDR.Close()
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            GetNewDocumentNumber = "N¤" & Err.Number & "»" & Err.Description
        Finally
            SqlCmd.Dispose()
        End Try
    End Function
    Private Function SaveData() As Boolean
        Try
            'STORED DETAILS VALUES 
            strYYYYmm = Year(ConvertToDate(DTPIMNT.Text)) & VB.Right("0" & Month(ConvertToDate(DTPIMNT.Text)), 2)
            Dim TEMP_INTRAMAT_DTL As New DataTable("TEMP_INTRAMAT_DTL")

            Dim IP_ADDRESS As New DataColumn("IP_ADDRESS_DTL", System.Type.GetType("System.String"))
            Dim Num_Id As New DataColumn("Num_Id", System.Type.GetType("System.Decimal"))
            Dim ITEM_CODE As New DataColumn("ITEM_CODE", System.Type.GetType("System.String"))
            Dim ITEM_GLGRP As New DataColumn("ITEM_GLGRP", System.Type.GetType("System.String"))
            Dim ITEM_GLCODE As New DataColumn("ITEM_GLCODE", System.Type.GetType("System.String"))
            Dim ITEM_SLCODE As New DataColumn("ITEM_SLCODE", System.Type.GetType("System.String"))
            Dim CUST_DRG_NO As New DataColumn("CUST_DRG_NO", System.Type.GetType("System.String"))
            Dim RATE As New DataColumn("RATE", System.Type.GetType("System.Decimal"))
            Dim QTY As New DataColumn("QTY", System.Type.GetType("System.Decimal"))
            Dim LOCATION_CODE As New DataColumn("LOCATION_CODE", System.Type.GetType("System.String"))
            Dim UNIT_CODE As New DataColumn("UNIT_CODE_DTL", System.Type.GetType("System.String"))

            TEMP_INTRAMAT_DTL.Columns.Add(IP_ADDRESS)
            TEMP_INTRAMAT_DTL.Columns.Add(Num_Id)
            TEMP_INTRAMAT_DTL.Columns.Add(ITEM_CODE)
            TEMP_INTRAMAT_DTL.Columns.Add(ITEM_GLGRP)
            TEMP_INTRAMAT_DTL.Columns.Add(ITEM_GLCODE)
            TEMP_INTRAMAT_DTL.Columns.Add(ITEM_SLCODE)
            TEMP_INTRAMAT_DTL.Columns.Add(CUST_DRG_NO)
            TEMP_INTRAMAT_DTL.Columns.Add(RATE)
            TEMP_INTRAMAT_DTL.Columns.Add(QTY)
            TEMP_INTRAMAT_DTL.Columns.Add(LOCATION_CODE)
            TEMP_INTRAMAT_DTL.Columns.Add(UNIT_CODE)

            With Me.SprIMNTGrid
                For intLoopCounter = 1 To .MaxRows
                    VAR_ITEMCODE = Nothing
                    .GetText(ENUMDETAILS.VAL_ITEMCODE, intLoopCounter, VAR_ITEMCODE)
                    VAR_ITEMGRPCODE = Nothing
                    .GetText(ENUMDETAILS.VAL_ITEMGRPCODE, intLoopCounter, VAR_ITEMGRPCODE)
                    VAR_ITEMGLCODE = Nothing
                    .GetText(ENUMDETAILS.VAL_ITEMGLCODE, intLoopCounter, VAR_ITEMGLCODE)
                    VAR_ITEMSLCODE = Nothing
                    .GetText(ENUMDETAILS.VAL_ITEMSLCODE, intLoopCounter, VAR_ITEMSLCODE)
                    VAR_CUSTPARTNO = Nothing
                    .GetText(ENUMDETAILS.VAL_CUSTPARTNO, intLoopCounter, VAR_CUSTPARTNO)
                    VAR_RATE = Nothing
                    .GetText(ENUMDETAILS.VAL_RATE, intLoopCounter, VAR_RATE)
                    VAR_QUANTITY = Nothing
                    .GetText(ENUMDETAILS.VAL_QUANTITY, intLoopCounter, VAR_QUANTITY)
                    VAR_LOCATION_CODE = Nothing
                    .GetText(ENUMDETAILS.VAL_LOCATION, intLoopCounter, VAR_LOCATION_CODE)

                    TEMP_INTRAMAT_DTL.Rows.Add(gstrIpaddressWinSck.ToString.Trim, Convert.ToInt16(intLoopCounter), VAR_ITEMCODE.ToString.Trim, VAR_ITEMGRPCODE.ToString.Trim, VAR_ITEMGLCODE.ToString.Trim, VAR_ITEMSLCODE.ToString.Trim, VAR_CUSTPARTNO.ToString.Trim, Convert.ToDecimal(VAR_RATE), Convert.ToInt32(VAR_QUANTITY), VAR_LOCATION_CODE.ToString.Trim, gstrUNITID.ToString.Trim)
                Next
            End With
            If TEMP_INTRAMAT_DTL.Rows.Count <= 0 Then
                MessageBox.Show("Please select all the related information first.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            'END HERE

            'STORED ALL THE TRANSACTIONS
            SqlCmd = New SqlCommand
            Try
                
                With SqlCmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_SAVE_INTRAMATERIALTRANSFERNOTE"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID.ToString.Trim
                    .Parameters.Add("@RECV_UNIT_CODE", SqlDbType.VarChar, 10).Value = lblReceivingUnitCode.Text
                    .Parameters.Add("@YYYYMM", SqlDbType.Int).Value = Convert.ToInt64(strYYYYmm)
                    .Parameters.Add("@Customer_Code", SqlDbType.VarChar, 10).Value = txtCustomerCode.Text.ToString.Trim
                    .Parameters.Add("@Account_Ledger", SqlDbType.VarChar, 12).Value = mStrGLCode.ToString.Trim
                    .Parameters.Add("@Account_SubLedger", SqlDbType.VarChar, 12).Value = mStrSLCode.ToString.Trim
                    .Parameters.Add("@Cust_Ref", SqlDbType.VarChar, 25).Value = txtSoNo.Text.ToString.Trim
                    .Parameters.Add("@Amendment_No", SqlDbType.VarChar, 25).Value = lblAmendementNo.Text.ToString.Trim
                    .Parameters.Add("@CreditTerms", SqlDbType.VarChar, 4).Value = mStrCreditTerms.ToString.Trim
                    .Parameters.Add("@BasicDueDate", SqlDbType.Date, 11).Value = DtpBasicDueDate.Value
                    .Parameters.Add("@BasicPayDueDate", SqlDbType.Date, 11).Value = DtpBasicPayDate.Value
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 16).Value = mP_User.ToString.Trim
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck.ToString.Trim
                    Dim TblValue_DTL As New SqlParameter("@TEMP_TABLE_DTL", SqlDbType.Structured)
                    TblValue_DTL.TypeName = "TEMP_INTRAMATRIAL_DTL"
                    TblValue_DTL.Value = TEMP_INTRAMAT_DTL
                    .Parameters.Add(TblValue_DTL)
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 500).Value = ""
                    .Parameters("@MSG").Direction = ParameterDirection.InputOutput
                    .Parameters.Add("@DOC_NO", SqlDbType.Int).Value = 0
                    .Parameters("@DOC_NO").Direction = ParameterDirection.InputOutput
                    SqlCmd.ExecuteNonQuery()
                    If .Parameters("@MSG").Value.ToString.Trim.Length = 0 Then
                        txtIMNTNo.Text = .Parameters("@Doc_No").Value.ToString
                        mStrMSG = "Intra Material Transfer Note has been generated with " & txtIMNTNo.Text.ToString.Trim
                    Else
                        mStrMSG = .Parameters("@MSG").Value.ToString.Trim
                        Return False
                    End If
                End With
                Return True
            Catch SqlEx As SqlException
                Dim myError As SqlError
                MessageBox.Show("Errors Count:" & SqlEx.Errors.Count, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                For Each myError In SqlEx.Errors
                    MessageBox.Show(myError.Number & " - " & myError.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Next
                Return False
            Finally
                If SqlCmd.Connection.State = ConnectionState.Open Then
                    SqlCmd.Connection.Close()
                End If
            End Try

            'END HERE
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
        End Try

    End Function
    Private Function Validate_data() As Boolean
        Try
            Dim intCtr As Integer = 0
            If lblReceivingUnitCode.Text.Trim.Length = 0 Then
                'MsgBox("Please select To Unit first.", MsgBoxStyle.Information, ResolveResString(100))
                mStrMSG = "Please select To Unit first."
                Return False
            End If
            mStrMSG = ""
            For intCtr = 1 To SprIMNTGrid.MaxRows
                With Me.SprIMNTGrid
                    VAR_ITEMCODE = Nothing
                    .GetText(ENUMDETAILS.VAL_ITEMCODE, intCtr, VAR_ITEMCODE)
                    VAR_CUSTPARTNO = Nothing
                    .GetText(ENUMDETAILS.VAL_CUSTPARTNO, intCtr, VAR_CUSTPARTNO)
                    VAR_QUANTITY = Nothing
                    .GetText(ENUMDETAILS.VAL_QUANTITY, intCtr, VAR_QUANTITY)
                End With
                If Convert.ToDecimal(VAR_QUANTITY) = 0 Then
                    mStrMSG = "Quantity can not be 0 for item code:- " & VAR_ITEMCODE.ToString.Trim
                    Return False
                End If
                CheckcustorddtlQty(VAR_ITEMCODE.ToString.Trim, VAR_CUSTPARTNO.ToString.Trim, Convert.ToDecimal(VAR_QUANTITY))
            Next
            If mStrMSG <> "" Then
                Return False
            End If
            mStrMSG = ""
            If Not CheckforSchedules() = "" Then
                Return False
            End If
            mStrMSG = ""
            If CheckBalanceAT_Location() = False Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Function
    Public Function CheckcustorddtlQty(ByRef pstrItemCode As String, ByRef pstrDrgno As String, ByRef pdblQty As Double) As String
        Try
            mStrSql = "Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & _
            " Account_code ='" & txtCustomerCode.Text & "'" & " and Item_code ='" & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno & "' " & _
            " and Authorized_flag = 1 and cust_ref = '" & txtSoNo.Text & "' and Amendment_no = '" & lblAmendementNo.Text & "'"

            SqlCmd = New SqlCommand
            With SqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = mStrSql
                SqlDR = SqlCmd.ExecuteReader()
                If SqlDR.Read = True Then
                    If SqlDR("OpenSO") = True Then
                        Return ""
                    Else
                        If Val(SqlDR("Balance_Qty")) < pdblQty Then
                            mStrMSG = mStrMSG & " Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & Val(SqlDR("Balance_Qty")) & "." & vbCrLf
                            'MessageBox.Show("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & Val(SqlDR("Balance_Qty")) & ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            Return mStrMSG
                        End If
                    End If
                End If
            End With
            If SqlDR.IsClosed = False Then SqlDR.Close()
            SqlCmd.Connection.Close()
            SqlCmd.Dispose()
            Return mStrMSG
        Catch ex As Exception
            mStrMSG = "Error"
            Return mStrMSG
        End Try
    End Function
    Private Function CheckforSchedules() As String
        Try

            Dim intCtr As Integer
            Dim intCount As Integer = 0
            Dim intLoop As Integer = 0
            ReDim mSchTypeArr(0)
            CheckforSchedules = ""
            strYYYYmm = Year(ConvertToDate(DTPIMNT.Text)) & VB.Right("0" & Month(ConvertToDate(DTPIMNT.Text)), 2)
            intCount = 0
            intLoop = 0
            For intCtr = 1 To SprIMNTGrid.MaxRows
                With Me.SprIMNTGrid
                    VAR_ITEMCODE = Nothing
                    .GetText(ENUMDETAILS.VAL_ITEMCODE, intCtr, VAR_ITEMCODE)
                    VAR_CUSTPARTNO = Nothing
                    .GetText(ENUMDETAILS.VAL_CUSTPARTNO, intCtr, VAR_CUSTPARTNO)
                    VAR_QUANTITY = Nothing
                    .GetText(ENUMDETAILS.VAL_QUANTITY, intCtr, VAR_QUANTITY)
                End With

                intLoop = intLoop + 1
                intCount = intCount + 1
                'If StrComp(VAR_ITEMCODE, marrPrevQty(intLoop).ItemCode, vbTextCompare) = 0 And StrComp(VAR_CUSTPARTNO, marrPrevQty(intLoop).CustItemCode, vbTextCompare) = 0 Then
                '    VAR_QUANTITY = Val(VAR_QUANTITY) - marrPrevQty(intLoop).PrevQuantity
                'End If

                ReDim Preserve mSchTypeArr(intCount)
                SqlCmd = New SqlCommand
                With SqlCmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "MKT_SCHEDULE_CHECK_NORTH"
                    .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 10).Value = gstrUNITID.ToString.Trim
                    .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 8).Value = txtCustomerCode.Text.ToString.Trim
                    .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = VAR_ITEMCODE.ToString
                    .Parameters.Add("@CUSTDRG_NO", SqlDbType.VarChar, 30).Value = VAR_CUSTPARTNO.ToString.Trim
                    .Parameters.Add("@REQ_QTY", SqlDbType.Money).Value = VAR_QUANTITY
                    .Parameters.Add("@YYYYMM", SqlDbType.Int).Value = Convert.ToInt64(strYYYYmm)
                    .Parameters.Add("@DATE", SqlDbType.VarChar, 11).Value = getDateForDB(DTPIMNT.Text)
                    .Parameters.Add("@SCH_TYPE", SqlDbType.Char, 1).Value = 1
                    .Parameters("@SCH_TYPE").Direction = ParameterDirection.Output
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 500).Value = 500
                    .Parameters("@MSG").Direction = ParameterDirection.Output
                    .Parameters.Add("@ERR", SqlDbType.VarChar, 500).Value = 100
                    .Parameters("@ERR").Direction = ParameterDirection.Output
                    .ExecuteNonQuery()
                    If Len(.Parameters(9).Value) > 0 Then
                        MsgBox(.Parameters(9).Value, vbInformation + vbOKOnly, ResolveResString(100))
                        CheckforSchedules = "Error"
                        SqlCmd.Dispose()
                        SqlCmd = Nothing
                        Exit Function
                    End If
                    If Len(.Parameters(8).Value) > 0 Then
                        mStrMSG = mStrMSG & .Parameters(8).Value
                    End If
                End With

            Next intCtr
            Return mStrMSG

        Catch ex As Exception
            Return "Error in Schedule checking."
        End Try

    End Function
    Private Function PopulateData() As Boolean
        Dim intLoopcounter As Integer = 1
        Dim sqlAdp As SqlDataAdapter = Nothing
        Dim sqlDataSet As DataSet = Nothing
        Try
            Call InitializeSpreed()
            If CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                mStrSql = " select h.Doc_Date,h.Customer_Code,h.Cust_Ref,h.Amendment_No,h.AR_DrNote,h.AR_DrDate, " & _
                                        " h.Booked_FinanceValue,d.Item_Code,d.Cust_DrgNo,d.Item_rate,d.MTN_Qty,d.Location_Code, " & _
                                        " INV.invGld_invGLGrpId,INV.invGld_glCode,INV.invGld_slCode,C.CUST_NAME,CUST.Cust_Drg_Desc " & _
                                        " from IMTN_Hdr h inner join IMTN_Dtl d on h.Unit_Code = d.Unit_code and h.Doc_No = d.Doc_no  " & _
                                        " and h.Doc_Type = d.Doc_Type inner join item_mst i  " & _
                                        " on d.Item_Code = i.Item_Code and d.Unit_code = i.UNIT_CODE " & _
                                        " inner join fin_invGLGrpDtl inv on i.GlGrp_code = inv.invGld_invGLGrpId  " & _
                                        " and i.UNIT_CODE = inv.UNIT_CODE " & _
                                        " INNER JOIN CUSTOMER_MST C ON h.Customer_Code = C.CUSTOMER_CODE AND H.UNIT_CODE = C.UNIT_CODE " & _
                                        " INNER JOIN Cust_ord_dtl CUST ON D.ITEM_CODE = CUST.ITEM_CODE AND D.Cust_DrgNo = CUST.Cust_DrgNo " & _
                                        " AND H.CUSTOMER_CODE = CUST.ACCOUNT_CODE AND  h.Cust_Ref = cust.Cust_Ref AND H.UNIT_CODE = CUST.UNIT_CODE " & _
                                        " and h.Amendment_No = cust.Amendment_No " & _
                                        " where inv.invGld_prpsCode = 'StockTrans' " & _
                                        " AND H.DOC_NO = '" & txtIMNTNo.Text.ToString.Trim & "' " & _
                                        " AND H.UNIT_CODE = '" & gstrUNITID.ToString.Trim & "'"
            End If

            If CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdButtonGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                mStrSql = " select h.Doc_Date,h.Customer_Code,h.Cust_Ref,h.Amendment_No,h.AR_DrNote,h.AR_DrDate, " & _
                        " h.Booked_FinanceValue,d.Item_Code,d.Cust_DrgNo,d.Item_rate,d.MTN_Qty,d.Location_Code, " & _
                        " INV.invGld_invGLGrpId,INV.invGld_glCode,INV.invGld_slCode,C.CUST_NAME,CUST.Cust_Drg_Desc " & _
                        " from IMTN_Hdr h inner join IMTN_Dtl d on h.Unit_Code = d.Unit_code and h.Doc_No = d.Doc_no  " & _
                        " and h.Doc_Type = d.Doc_Type inner join item_mst i  " & _
                        " on d.Item_Code = i.Item_Code and d.Unit_code = i.UNIT_CODE " & _
                        " inner join fin_invGLGrpDtl inv on i.GlGrp_code = inv.invGld_invGLGrpId  " & _
                        " and i.UNIT_CODE = inv.UNIT_CODE " & _
                        " INNER JOIN CUSTOMER_MST C ON h.Customer_Code = C.CUSTOMER_CODE AND H.UNIT_CODE = C.UNIT_CODE " & _
                        " INNER JOIN Cust_ord_dtl CUST ON D.ITEM_CODE = CUST.ITEM_CODE AND D.Cust_DrgNo = CUST.Cust_DrgNo " & _
                        " AND H.CUSTOMER_CODE = CUST.ACCOUNT_CODE AND  h.Cust_Ref = cust.Cust_Ref AND H.UNIT_CODE = CUST.UNIT_CODE " & _
                        " and h.Amendment_No = cust.Amendment_No " & _
                        " where inv.invGld_prpsCode = 'StockTrans' " & _
                        " AND H.DOC_NO = '" & txtIMNTNo.Text.ToString.Trim & "' " & _
                        " AND H.UNIT_CODE = '" & gstrUNITID.ToString.Trim & "' and cust.Active_Flag = 'A'"
            End If
            sqlAdp = New SqlDataAdapter(mStrSql, SqlConnectionclass.GetConnection)
            sqlDataSet = New DataSet
            sqlAdp.Fill(sqlDataSet)
            With SprIMNTGrid
                If sqlDataSet.Tables(0).Rows.Count > 0 Then
                    DTPIMNT.Value = sqlDataSet.Tables(0).Rows(0)("Doc_Date").ToString
                    txtCustomerCode.Text = sqlDataSet.Tables(0).Rows(0)("Customer_Code").ToString
                    lblCustomerDesc.Text = sqlDataSet.Tables(0).Rows(0)("CUST_NAME").ToString
                    txtSoNo.Text = sqlDataSet.Tables(0).Rows(0)("Cust_Ref").ToString
                    lblAmendementNo.Text = sqlDataSet.Tables(0).Rows(0)("Amendment_No").ToString
                    lblARDrNote.Text = sqlDataSet.Tables(0).Rows(0)("AR_DrNote").ToString
                    For intRowCount = 0 To sqlDataSet.Tables(0).Rows.Count - 1
                        Call AddBlankRow()
                        .SetText(ENUMDETAILS.VAL_ITEMCODE, intLoopcounter, sqlDataSet.Tables(0).Rows(intRowCount)("Item_Code").ToString)
                        .SetText(ENUMDETAILS.VAL_CUSTPARTNO, intLoopcounter, sqlDataSet.Tables(0).Rows(intRowCount)("Cust_DrgNo").ToString)
                        .SetText(ENUMDETAILS.VAL_RATE, intLoopcounter, Convert.ToDecimal(sqlDataSet.Tables(0).Rows(intRowCount)("Item_rate")).ToString)
                        .SetText(ENUMDETAILS.VAL_QUANTITY, intLoopcounter, Convert.ToDecimal(sqlDataSet.Tables(0).Rows(intRowCount)("MTN_Qty")).ToString)
                        .SetText(ENUMDETAILS.VAL_CUSTPARTNODESC, intLoopcounter, sqlDataSet.Tables(0).Rows(intRowCount)("Cust_Drg_Desc").ToString())
                        .SetText(ENUMDETAILS.VAL_CURRSTOCK, intLoopcounter, Convert.ToString(GetStockInHand(sqlDataSet.Tables(0).Rows(intRowCount)("Location_Code").ToString, sqlDataSet.Tables(0).Rows(intRowCount)("Item_Code").ToString)))
                        .SetText(ENUMDETAILS.VAL_LOCATION, intLoopcounter, sqlDataSet.Tables(0).Rows(intRowCount)("Location_Code").ToString)
                        .SetText(ENUMDETAILS.VAL_ITEMGRPCODE, intLoopcounter, sqlDataSet.Tables(0).Rows(intRowCount)("invGld_invGLGrpId").ToString)
                        .SetText(ENUMDETAILS.VAL_ITEMGLCODE, intLoopcounter, sqlDataSet.Tables(0).Rows(intRowCount)("invGld_glCode").ToString)
                        .SetText(ENUMDETAILS.VAL_ITEMSLCODE, intLoopcounter, sqlDataSet.Tables(0).Rows(intRowCount)("invGld_SlCode").ToString)
                        intLoopcounter = intLoopcounter + 1
                        PopulateData = True
                    Next
                    SprIMNTGrid.Enabled = True
                    CmdButtonGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                Else
                    CmdButtonGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                    Return False
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
        End Try
    End Function
    Private Function CheckBalanceAT_Location() As Boolean
        Try
            With SprIMNTGrid
                For intLoopCounter = 1 To .MaxRows
                    VAR_ITEMCODE = Nothing
                    .GetText(ENUMDETAILS.VAL_ITEMCODE, intLoopCounter, VAR_ITEMCODE)
                    VAR_LOCATION_CODE = Nothing
                    .GetText(ENUMDETAILS.VAL_LOCATION, intLoopCounter, VAR_LOCATION_CODE)
                    VAR_QUANTITY = Nothing
                    .GetText(ENUMDETAILS.VAL_QUANTITY, intLoopCounter, VAR_QUANTITY)
                    mStrSql = " SELECT CUR_BAL FROM ITEMBAL_MST WHERE UNIT_CODE = '" & gstrUNITID.ToString.Trim & "' AND ITEM_CODE = '" & VAR_ITEMCODE.ToString.Trim & "' AND LOCATION_CODE = '" & VAR_LOCATION_CODE.ToString.Trim & "'"
                    SqlCmd = New SqlCommand
                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandType = CommandType.Text
                        .CommandText = mStrSql
                        SqlDR = SqlCmd.ExecuteReader()
                        If SqlDR.Read = True Then
                            If Convert.ToDecimal(VAR_QUANTITY) > Convert.ToDecimal(SqlDR("CUR_BAL")) Then
                                mStrMSG = mStrMSG & "Current balance is not avaliable at location " & VAR_LOCATION_CODE.ToString.Trim & " for Item Code " & VAR_ITEMCODE.ToString & vbCrLf
                            End If
                        End If
                        If SqlDR.IsClosed = False Then SqlDR.Close()
                        SqlCmd.Connection.Close()
                        SqlCmd.Dispose()
                    End With
                Next
            End With
            If mStrMSG.ToString.Trim.Length > 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        End Try
    End Function
    Private Function Fetchdata_rpt() As Boolean
        Try
            SqlCmd = New SqlCommand
            With SqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.StoredProcedure
                .CommandText = "Usp_imtnreports"
                .CommandTimeout = 0
                .Parameters.Add("@MODE", SqlDbType.VarChar, 10).Value = "T"
                .Parameters.Add("@IPADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck.ToString.Trim
                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID.ToString.Trim
                .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 15).Value = txtIMNTNo.Text.ToString.Trim
                .Parameters.Add("@Date_from", SqlDbType.Date, 11).Value = DTPIMNT.Value
                .Parameters.Add("@Date_to", SqlDbType.Date, 11).Value = DTPIMNT.Value
                .Parameters.Add("@Customer_Code", SqlDbType.VarChar, 10).Value = txtCustomerCode.Text.ToString.Trim
                SqlDR = SqlCmd.ExecuteReader()
                If SqlDR.HasRows = True Then
                    MessageBox.Show(SqlDR(0), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                    SqlDR.Close()
                    SqlCmd.Dispose()
                    Return False
                End If
            End With
            SqlDR.Close()
            SqlCmd.Dispose()
            Return True
        Catch SqlEx As SqlException
            Dim myError As SqlError
            MessageBox.Show("Errors Count:" & SqlEx.Errors.Count, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            For Each myError In SqlEx.Errors
                MessageBox.Show(myError.Number & " - " & myError.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Next
            SqlCmd.Dispose()
            Return False
        End Try
    End Function

#End Region

   


End Class