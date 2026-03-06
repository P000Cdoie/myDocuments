Option Strict Off
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports VB = Microsoft.VisualBasic
Imports System.Globalization
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic.FileIO

Friend Class frmMKTTRN0066
    Inherits System.Windows.Forms.Form
    '****************************************************
    'Copyright (c)  -  MIND
    'Name of module -  FRMMKTTRN0066.frm
    'Created By     -  Shubhra Verma
    'Created On     -  13 apr 2010
    'Issue ID       -  eMpro-20100416-45429
    'description    -  Trigger Uploading
    'Modified By Sanchi on 20 May 2011
    '   Modified to support MultiUnit functionality
    '============================================================================================
    'Revised By         :   Shabbir Hussain
    'Revised On         :   22 Dec 2011
    'Reason             :   Changes for Unit IV & FSP Trigger mapping
    '============================================================================================
    'MODIFIED BY NITIN MEHTA 0N 31 JAN 2012 FOR CHANGE MANAGEMENT
    '============================================================================================
    'Revised By         :   Shubhra Verma
    'Revised On         :   Nov 2015
    'ISSUE ID           :   10933955 
    '============================================================================================
    'Revised By         :   Shubhra Verma
    'Revised On         :   MAY 2015
    'ISSUE ID           :   10933955 
    '============================================================================================
    'Modified By         :   Milind Mishra
    'Modified On         :   22/11/2017
    'ISSUE ID            :   101406664 - Nissan auto invoice process enhancement
    '============================================================================================
    'Modified By         :   Praveen Kumar
    'Modified On         :   27/01/2022
    'ISSUE ID            :   10155819 — Nissan RAN file upload issue
    '============================================================================================

    Dim m_strCustomerCode As String
    Dim mintFormIndex As Short
    Dim StrDocNum As String
    Dim mbln_View As Boolean
    Dim Flag As Boolean
    Private Enum enm_mapping
        DrgNo = 1
        Desc
        Format
        SNPFormat
        ReceivingArea
        SupplyArea
    End Enum

    Private Enum enm_SNPmapping
        DrgNo = 1
        Desc
        SNPFormat

    End Enum
    Private Sub ctlFormHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader.Click
        Try
            Call ShowHelp("UNDERCONSTRUCTION.HTM")
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

#Region "Form Events"
    Private Sub frmMKTTRN0058_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        Try
            'Form number assigned for the current form
            mdifrmMain.CheckFormName = mintFormIndex
            'Form name text is made BOLD
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub frmMKTTRN0058_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Try
            'Form name would be adjusted to normal
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
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
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub frmMKTTRN0058_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try
            Call FitToClient(Me, frmMain, ctlFormHeader, pnlmain, 500)
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.HeaderString())
            Call EnableControls(True, Me)
            lblFileName.BackColor = Me.BackColor
            cmdCustHelp.Enabled = False
            txtCustomerCode.Enabled = False
            SetSpreadProperty()
            'SetSNPSpreadProperty()
        Catch ex As Exception
            RaiseException(ex)
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
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
#End Region

#Region "Upload Triggers"
    Private Sub CmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdBrowse.Click

        Try
            CommanDLogOpen.InitialDirectory = gstrLocalCDrive
            CommanDLogOpen.Filter = "Microsoft Excel File  (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv"
            CommanDLogOpen.ShowDialog()
            lblFileName.Text = CommanDLogOpen.FileName
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub CmdUploadFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdUploadFile.Click
        '10933955
        Try
            Dim FileName As String = Nothing

            If optNissan.Checked = True Then
                If txtCustomerCode.Text.Trim.Length = 0 Then
                    MessageBox.Show("Select Customer.", ResolveResString(100), MessageBoxButtons.OK)
                    If cmdCustHelp.Enabled = True Then cmdCustHelp.Focus()
                    Exit Sub
                End If
                If lblFileName.Text.Trim.Length = 0 Then
                    MessageBox.Show("First select file to be uploaded !", ResolveResString(100), MessageBoxButtons.OK)
                    If CmdBrowse.Enabled = True Then CmdBrowse.Focus()
                    Exit Sub
                End If
            End If

            If optPCA.Checked = True Then
                If txtCustomerCode.Text.Trim.Length = 0 Then
                    MessageBox.Show("Please select Customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    If cmdCustHelp.Enabled = True Then cmdCustHelp.Focus()
                    Exit Sub
                End If
                If lblFileName.Text.Trim.Length = 0 Then
                    MessageBox.Show("First select file to be uploaded !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    If CmdBrowse.Enabled = True Then CmdBrowse.Focus()
                    Exit Sub
                End If
            End If

            If lblFileName.Text.Trim.Length = 0 Then
                MessageBox.Show("First select file to be uploaded !", ResolveResString(100), MessageBoxButtons.OK)
                If CmdBrowse.Enabled = True Then CmdBrowse.Focus()
                Exit Sub
            End If

            FileName = lblFileName.Text.Trim.Substring(lblFileName.Text.Trim.LastIndexOf("\") + 1, lblFileName.Text.Trim.Length - lblFileName.Text.Trim.LastIndexOf("\") - 1)

            Dim DT As DataTable = SqlConnectionclass.GetDataTable("select top 1 Auto_id  from Ford_Trigger_File  where Uploaded_FileName = '" & FileName.ToString().Trim() & "'")
            If DT.Rows.Count > 0 Then
                MessageBox.Show("File already uploaded, select another file ")
                Exit Sub
            End If



            If optmanual.Checked = True Then
                GETDataFromExcelManual(lblFileName.Text.Trim)
            ElseIf opteilvs.Checked = True Then
                GETDataFromExceleILVS(lblFileName.Text.Trim)
            ElseIf optNissan.Checked = True Then
                GETDataFromExcel_NISSAN(lblFileName.Text.Trim)
            ElseIf optPCA.Checked = True Then
                GETDataFromTextFile_PCA(lblFileName.Text.Trim)
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally

        End Try

    End Sub

    Private Sub GETDataFromExceleILVS(ByVal filepath As String)

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim range As Excel.Range
        Dim rCnt As Integer
        Dim strtime As String = ""

        Dim strReaderdate As Object : Dim strReadertime As Object
        Dim strRotationNbr As Object : Dim strBlendNumber As Object
        Dim strProdSys As Object

        Dim strVin As Object : Dim strVihicleLine As Object
        Dim strOfflineDate As Object : Dim strEvent As Object
        Dim strLaunchPrg As Object

        Dim strBuildLevelDate As Object : Dim strLineFeed As Object
        Dim strPartBase As Object : Dim strPartPreFix As Object
        Dim strPartSuffix As Object

        Dim strDescription As Object : Dim strUsages_qty As Object
        Dim strSupplier As Object : Dim strCPSC As Object
        Dim strVehicleSeqNo As Object
        Dim strInstallArea As Object
        Dim ss As DateTime
        Dim strsql As String = ""
        Dim IsCompleted As Boolean = False

        Try
            xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Open(filepath)
            xlWorkSheet = xlWorkBook.Worksheets(1)
            range = xlWorkSheet.UsedRange

            PB.Visible = True
            PB.Minimum = 0
            PB.Maximum = range.Rows.Count + 20
            Dim FileName As String = Nothing

            FileName = filepath.Substring(filepath.LastIndexOf("\") + 1, filepath.Length - filepath.LastIndexOf("\") - 1)

            SqlConnectionclass.BeginTrans()

            For rCnt = 1 To range.Rows.Count
                PB.Value = rCnt
                If rCnt <> 1 Then
                    strReaderdate = CType(range.Cells(rCnt, 1), Excel.Range)
                    strReadertime = CType(range.Cells(rCnt, 2), Excel.Range)
                End If

                strRotationNbr = CType(range.Cells(rCnt, 3), Excel.Range)
                strBlendNumber = CType(range.Cells(rCnt, 4), Excel.Range)
                strProdSys = CType(range.Cells(rCnt, 5), Excel.Range)

                strVin = CType(range.Cells(rCnt, 6), Excel.Range)
                strVihicleLine = CType(range.Cells(rCnt, 7), Excel.Range)
                strOfflineDate = CType(range.Cells(rCnt, 8), Excel.Range)
                strEvent = CType(range.Cells(rCnt, 9), Excel.Range)
                strLaunchPrg = CType(range.Cells(rCnt, 10), Excel.Range)

                strBuildLevelDate = CType(range.Cells(rCnt, 11), Excel.Range)
                strLineFeed = CType(range.Cells(rCnt, 12), Excel.Range)
                strPartBase = CType(range.Cells(rCnt, 13), Excel.Range)
                strPartPreFix = CType(range.Cells(rCnt, 14), Excel.Range)
                strPartSuffix = CType(range.Cells(rCnt, 15), Excel.Range)

                strDescription = CType(range.Cells(rCnt, 16), Excel.Range)
                strUsages_qty = CType(range.Cells(rCnt, 17), Excel.Range)
                strSupplier = CType(range.Cells(rCnt, 18), Excel.Range)
                strCPSC = CType(range.Cells(rCnt, 19), Excel.Range)
                strVehicleSeqNo = CType(range.Cells(rCnt, 20), Excel.Range)
                strInstallArea = CType(range.Cells(rCnt, 21), Excel.Range)

                If rCnt = 1 Then
                    If strRotationNbr.Value.Trim.ToUpper() <> "Rotation Nbr".ToUpper() Or strBlendNumber.Value.Trim.ToUpper() <> "Blend Number".ToUpper() Or strProdSys.Value.Trim.ToUpper() <> "Prod Sys".ToUpper() Then
                        MessageBox.Show("Invalid File Format " & vbCrLf & " File Format " & vbCrLf & "  " & _
                                        " [Reader Date|Reader Time|Rotation Nbr|Blend Number|Prod Sys] " & vbCrLf & " " & _
                                        " [VIN|Vehicle Line|Offline Date|Event|Launch Pgm|Build Level Date] " & vbCrLf & "  " & _
                                        " [Line Feed|Part Base|Part Prefix|Part Sufix|Description|Usage Qty] " & vbCrLf & "  " & _
                                        " [Supplier|CPSC|Vehicle Sequence No|Install Area] ")
                        Exit For
                    End If

                    If strVin.Value.Trim.ToUpper() <> "VIN".ToUpper() Or strVihicleLine.Value.Trim.ToUpper() <> "Vehicle Line".ToUpper() Or strOfflineDate.Value.Trim.ToUpper() <> "Offline Date".ToUpper() Or strEvent.Value.Trim.ToUpper() <> "Event".ToUpper() Or strLaunchPrg.Value.Trim.ToUpper() <> "Launch Pgm".ToUpper() Or strBuildLevelDate.Value.Trim.ToUpper() <> "Build Level Date".ToUpper() Or strLineFeed.Value.Trim.ToUpper() <> "Line Feed".ToUpper() Or strPartBase.Value.Trim.ToUpper() <> "Part Base".ToUpper() Or strPartPreFix.Value.Trim.ToUpper() <> "Part Prefix".ToUpper() Or strPartSuffix.Value.Trim.ToUpper() <> "Part Sufix".ToUpper() Or strDescription.Value.Trim.ToUpper() <> "Description".ToUpper() Or strUsages_qty.Value.Trim.ToUpper() <> "Usage Qty".ToUpper() Or strSupplier.Value.Trim.ToUpper() <> "Supplier".ToUpper() Or strCPSC.Value.Trim.ToUpper() <> "CPSC".ToUpper() Or strVehicleSeqNo.Value.Trim.ToUpper() <> "Vehicle Sequence No".ToUpper() Or strInstallArea.Value.Trim.ToUpper() <> "Install Area".ToUpper() Then
                        MessageBox.Show("Invalid File Format " & vbCrLf & " File Format " & vbCrLf & "  " & _
                                     " [Reader Date|Reader Time|Rotation Nbr|Blend Number|Prod Sys] " & vbCrLf & " " & _
                                     " [VIN|Vehicle Line|Offline Date|Event|Launch Pgm|Build Level Date] " & vbCrLf & "  " & _
                                     " [Line Feed|Part Base|Part Prefix|Part Sufix|Description|Usage Qty] " & vbCrLf & "  " & _
                                     " [Supplier|CPSC|Vehicle Sequence No|Install Area] ")
                        Exit For
                    End If

                End If

                If rCnt <> 1 Then
                    strsql = "INSERT INTO ford_trigger_file(VinNo,SeqNo,PartNo,Seq_date,Entered_date,Unit_Code,Upload_Source,Uploaded_FileName) "
                    strsql = strsql + "VALUES('" + strVin.Value.ToString() + "','" + strRotationNbr.Value.ToString() + "','" + strPartBase.Value.ToString() + "-" + strPartPreFix.Value.ToString() + "-" + strPartSuffix.Value.ToString() + "',replace(convert(varchar(20),cast('" + strReaderdate.Value.ToString() + "' as date),102),'.','')+'" + " " + strReadertime.Value.ToString().Replace(":", "") + "', "
                    strsql = strsql + "Getdate(),'" & gstrUNITID & "','eILVS','" & FileName & "')"

                    SqlConnectionclass.ExecuteNonQuery(strsql)

                    strsql = "INSERT INTO eILVS_ford_trigger_file(Reader_Date,Reader_Time,Rotation_Nbr,Blend_Number,Prod_Sys, " & _
                             " VIN,Vehicle_Line,Offline_Date,Event1,Launch_Pgm,Build_Level_Date,Line_Feed,Part_Base, " & _
                             " Part_Prefix,Part_Sufix,Description1,Usage_Qty,Supplier,CPSC,Vehicle_Sequence_No,Install_Area,Entry_dt) "
                    strsql = strsql + "VALUES('" + strReaderdate.Value.ToString() + "','" + strReadertime.Value.ToString() + "','" + strRotationNbr.Value.ToString() + "','" + strBlendNumber.Value.ToString() + "', '" + strProdSys.Value.ToString() + "', "
                    strsql = strsql + "'" + strVin.Value.ToString() + "','" + strVihicleLine.Value.ToString() + "','" + strOfflineDate.Value.ToString() + "','" + strEvent.Value.ToString() + "', '" + strLaunchPrg.Value.ToString() + "', "
                    strsql = strsql + "'" + strBuildLevelDate.Value.ToString() + "','" + strLineFeed.Value.ToString() + "','" + strPartBase.Value.ToString() + "','" + strPartPreFix.Value.ToString() + "', '" + strPartSuffix.Value.ToString() + "', "
                    strsql = strsql + "'" + strDescription.Value.ToString() + "','" + strUsages_qty.Value.ToString() + "','" + strSupplier.Value.ToString() + "','" + strCPSC.Value.ToString() + "', '" + strVehicleSeqNo.Value.ToString() + "','" + strInstallArea.Value.ToString() + "', "
                    strsql = strsql + "Getdate())"

                    SqlConnectionclass.ExecuteNonQuery(strsql)
                    IsCompleted = True
                End If
            Next

            SqlConnectionclass.CommitTran()
            PB.Maximum = 100
            xlWorkBook.Close()
            'xlApp.Quit()

            'releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
            If IsCompleted = True Then
                MessageBox.Show("Trigger file Uploaded Successfully !", ResolveResString(100), MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            PB.Maximum = 100
            PB.Visible = False
            KillExcelProcess(xlApp)
        End Try

    End Sub

    Private Sub GETDataFromExcelManual(ByVal filepath As String)

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim range As Excel.Range
        Dim rCnt As Integer
        Dim Objseqno As Object
        Dim ObjPartno As Object
        Dim ObjSeqDate As Object
        Dim ObjSeqTime As Object
        Dim ObjVinNo As Object
        Dim strsql As String = ""
        Dim Iscompleted As Boolean = False
        Dim FileName As String = ""

        Try
            xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Open(filepath)
            xlWorkSheet = xlWorkBook.Worksheets(1)
            FileName = filepath.Substring(filepath.LastIndexOf("\") + 1, filepath.Length - filepath.LastIndexOf("\") - 1)
            range = xlWorkSheet.UsedRange
            PB.Visible = True
            PB.Minimum = 0
            PB.Maximum = range.Rows.Count + 20

            SqlConnectionclass.BeginTrans()

            For rCnt = 1 To range.Rows.Count
                PB.Value = rCnt

                Objseqno = CType(range.Cells(rCnt, 1), Excel.Range)
                ObjPartno = CType(range.Cells(rCnt, 2), Excel.Range)
                ObjSeqDate = CType(range.Cells(rCnt, 3), Excel.Range)
                ObjSeqTime = CType(range.Cells(rCnt, 4), Excel.Range)
                ObjVinNo = CType(range.Cells(rCnt, 5), Excel.Range)

                If rCnt = 1 Then
                    If Objseqno.Value.Trim().ToUpper() <> "SeqNo".ToUpper() Or ObjPartno.Value.Trim.ToUpper() <> "PartNo".ToUpper() Or ObjSeqDate.Value.Trim.ToUpper() <> "Seq_date".ToUpper() Or ObjSeqTime.Value.Trim.ToUpper() <> "Seq_Time".ToUpper() Or ObjVinNo.Value.Trim.ToUpper() <> "VinNo".ToUpper() Then
                        MessageBox.Show("Invalid File Format " & vbCrLf & " File Format  [SeqNo|PartNo|Seq_date|Seq_Time|VinNo] " & vbCrLf & " Seq_date Format YYYYMMDD  " & vbCrLf & " Seq_Time Format  HHMMSS ")
                        Exit For
                    End If
                End If

                If rCnt <> 1 Then
                    strsql = "INSERT INTO ford_trigger_file(VinNo,SeqNo,PartNo,Seq_date,Entered_date,Unit_Code,Uploaded_FileName,Upload_Source) "
                    strsql = strsql + "VALUES('" + ObjVinNo.value.ToString() + "','" + Objseqno.value.ToString() + "','" + ObjPartno.value.ToString() + "','" + ObjSeqDate.value.ToString() + " " + ObjSeqTime.value.ToString() + "', "
                    strsql = strsql + "Getdate(),'" & gstrUNITID & "','" & FileName & "','Manual')"

                    SqlConnectionclass.ExecuteNonQuery(strsql)
                    Iscompleted = True
                End If

            Next

            SqlConnectionclass.CommitTran()
            PB.Maximum = 100
            xlWorkBook.Close()
            'xlApp.Quit()

            'releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            If Iscompleted = True Then
                MessageBox.Show("Data Upload Successfully")
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            PB.Maximum = 100
            PB.Visible = False
            KillExcelProcess(xlApp)
        End Try

    End Sub

    Private Sub GETDataFromExcel_NISSAN(ByVal filepath As String)
        '10933955
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim range As Excel.Range
        Dim rCnt As Integer
        Dim strtime As String = ""

        Dim intRANNo As Integer = 0

        Dim strSTA As Object : Dim strSTATUS As Object : Dim strDEPOT As Object
        Dim strPARTNO As Object : Dim strDUEDATE As Object
        Dim strDUETIME As Object
        Dim strDUEDATEstring As String
        Dim strISSUENO As Object : Dim strDIV As Object
        Dim strQTY As Object : Dim strDELPLACE As Object
        Dim strSNP As Object : Dim strPKGCODE As Object
        'Issue ID-10155819
        Dim strSHIPPEDQTY As Object : Dim strRECEIVEDQTY As Object : Dim strTRIPNO As Object

        Dim strBuildLevelDate As Object : Dim strLineFeed As Object
        Dim strPartBase As Object : Dim strPartPreFix As Object
        Dim strPartSuffix As Object

        Dim strDescription As Object : Dim strUsages_qty As Object
        Dim strSupplier As Object : Dim strCPSC As Object
        Dim strVehicleSeqNo As Object
        Dim strInstallArea As Object
        Dim ss As DateTime
        Dim strsql As String = ""
        Dim OBJCURDATE As Object = Nothing
        Dim provider As CultureInfo = CultureInfo.InvariantCulture
        Dim IsCompleted As Boolean = False
        Dim oRdr As SqlDataReader = Nothing
        Dim strException As String = String.Empty

        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(filepath)
            xlWorkSheet = xlWorkBook.Worksheets(1)
            range = xlWorkSheet.UsedRange

            PB.Visible = True
            PB.Minimum = 0
            PB.Maximum = range.Rows.Count + 20
            Dim FileName As String = Nothing

            SqlConnectionclass.ExecuteNonQuery("Set dateformat 'DMY'")

            FileName = filepath.Substring(filepath.LastIndexOf("\") + 1, filepath.Length - filepath.LastIndexOf("\") - 1)

            intRANNo = SqlConnectionclass.ExecuteScalar("Select ISNULL(AutoRANSeqNo,0) AutoRANSeqNo from SYSFILE " & _
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' and cust_code = '" & txtCustomerCode.Text & "'")

            SqlConnectionclass.BeginTrans()
            'OBJCURDATE = CStr(GetServerDate.Year) + Mid(CStr(100 + GetServerDate.Month), 2, 2) + Mid(CStr(100 + GetServerDate.Day), 2, 2)
            For rCnt = 1 To range.Rows.Count
                strDUEDATE = CType(range.Cells(rCnt, 5), Excel.Range)
                PB.Value = rCnt

                If rCnt <> 1 Then
                    'If CStr(CDate(strDUEDATE.VALUE.ToString).Year) + Mid(CStr(100 + CDate(strDUEDATE.VALUE.ToString).Month), 2, 2) + Mid(CStr(100 + CDate(strDUEDATE.VALUE.ToString).Day), 2, 2) < OBJCURDATE Then
                    '    Continue For
                    'End If
                End If

                strSTA = Nothing
                strSTATUS = Nothing
                strSTA = CType(range.Cells(rCnt, 1), Excel.Range)
                strSTATUS = CType(range.Cells(rCnt, 2), Excel.Range)

                strDEPOT = CType(range.Cells(rCnt, 3), Excel.Range)
                strPARTNO = CType(range.Cells(rCnt, 4), Excel.Range)

                strDUETIME = CType(range.Cells(rCnt, 6), Excel.Range)
                strISSUENO = CType(range.Cells(rCnt, 7), Excel.Range)
                strDIV = CType(range.Cells(rCnt, 8), Excel.Range)
                strQTY = CType(range.Cells(rCnt, 9), Excel.Range)
                'Issue ID-10155819
                strSHIPPEDQTY = CType(range.Cells(rCnt, 10), Excel.Range)
                strRECEIVEDQTY = CType(range.Cells(rCnt, 11), Excel.Range)
                strDELPLACE = CType(range.Cells(rCnt, 12), Excel.Range)

                strSNP = CType(range.Cells(rCnt, 13), Excel.Range)
                strPKGCODE = CType(range.Cells(rCnt, 14), Excel.Range)
                strTRIPNO = CType(range.Cells(rCnt, 15), Excel.Range)


                If rCnt = 1 Then
                    If strSTA.VALUE.ToString.ToUpper <> "STA" Or strSTATUS.VALUE.ToString.ToUpper <> "STATUS" Or strDEPOT.VALUE.ToString.ToUpper <> "DEPOT" _
                        Or strPARTNO.VALUE.ToString.ToUpper <> "PARTS NO (P/#)" Or strDUEDATE.VALUE.ToString.ToUpper <> "DUE DATE YYMMDD" _
                        Or strDUETIME.VALUE.ToString.ToUpper <> "DUE DATE HHMM" Or strISSUENO.value.ToString.ToUpper <> "RAN NO" _
                        Or strDIV.value.ToString.ToUpper <> "DIV" Or strQTY.value.ToString.ToUpper <> "ORDER QTY" Or strDELPLACE.value.ToString.ToUpper <> "DELIVERY PLACE" _
                        Or strSHIPPEDQTY.value.ToString.ToUpper <> "SHIPPED QTY" Or strRECEIVEDQTY.value.ToString.ToUpper <> "RECEIVED QTY" Or strTRIPNO.value.ToString.ToUpper <> "TRIP NO" _
                        Or strSNP.VALUE.ToString.ToUpper <> "SNP" Or strPKGCODE.VALUE.ToString.ToUpper <> "PACKING CODE" Then
                        ''Issue ID-10155819 --Changes in IF Condition
                        MessageBox.Show("Invalid File Format " & vbCrLf & " File Format " & vbCrLf & "  " & _
                                       " [STA|STATUS|DEPOT|Parts NO (P/#)|Due Date YYMMDD] " & vbCrLf & " " & _
                                       " [Due Date HHMM|RAN NO|DIV|ORDER QTY|SHIPPED QTY|RECEIVED QTY|DELIVERY PLACE|SNP|PACKING CODE|TRIP NO] ")
                        Exit For
                    End If
                End If

                If rCnt <> 1 Then

                    If isOADate(strDUEDATE.value.ToString) = "Y" Then
                        strDUEDATE = DateTime.FromOADate(strDUEDATE.Value).ToString("dd MMM yy")
                    Else
                        If isOAString(strDUEDATE.value.ToString) = "Y" Then
                            strDUEDATE = strDUEDATE.Value
                            strDUEDATEstring = strDUEDATE
                            strDUEDATEstring = DateTime.ParseExact(strDUEDATEstring, "dd/MM/yyyy", Nothing).ToString("MM/dd/yyyy")
                            'strDUEDATE = Convert.ToDateTime(strDUEDATE.Value).ToString("MM/dd/yyyy")
                            'strDUEDATE = DateTime.ParseExact(strDUEDATE.Value, "dd/MM/yyyy", provider)
                        Else
                            strDUEDATE = Date.ParseExact(strDUEDATE.value.ToString, "d", provider)
                        End If
                    End If

                    'strDUEDATE = CStr(CDate(strDUEDATE).Year) + Mid(CStr(100 + CDate(strDUEDATE).Month), 2, 2) + Mid(CStr(100 + CDate(strDUEDATE).Day), 2, 2)
                    strDUETIME = Val(strDUETIME.VALUE.ToString) * 24
                    'strDUETIME = Mid(100 + strDUETIME, 2, 2) + ":" + Mid(100 + ((strDUETIME - Int(strDUETIME)) * 60), 2, 2) + ":00"
                    strDUETIME = Mid(100 + strDUETIME, 2, 2) + ":" + Mid(100 + Int(((100 + strDUETIME) - Int(100 + strDUETIME)) * 60), 2, 2) + ":00"
                End If

                If rCnt <> 1 Then
                    If intRANNo >= 9999 Then
                        intRANNo = 0
                    End If
                    intRANNo = intRANNo + 1
                    ''Issue ID-10155819 --Insert Query Changes
                    strsql = "Set DateFormat 'DMY' INSERT INTO ford_trigger_file_NISSAN(VinNo,SeqNo,PartNo,Unit_Code,Seq_date," & _
                        " Entered_date,UPLOAD_SOURCE,Uploaded_FileName," & _
                        " STA,Status,Depot,DIV,QTY,DelPlace,SNP,PkgCode,CUST_CODE,IP_ADDRESS,shipped_qty,Received_Qty,Trip_no) " & _
                        " VALUES('" & strISSUENO.Value.ToString() & "'," & intRANNo & ",'" & strPARTNO.Value.ToString() & "','" & gstrUNITID & "',"
                    strsql = strsql & " Convert(CHAR(8),CONVERT(DATE,'" & strDUEDATEstring & "',101),112)  + ' ' + '" & strDUETIME & "', "
                    strsql = strsql & " Getdate(),'NISSAN','" & lblFileName.Text & "','" & strSTA.VALUE.ToString() & "','" & strSTATUS.VALUE.ToString() & "'," & _
                        " '" & strDEPOT.VALUE.ToString & "','" & IIf(IsNothing(strDIV.VALUE), "", strDIV.VALUE).ToString & "'," & _
                        "'" & strQTY.VALUE.ToString & "','" & strDELPLACE.VALUE.ToString & "','" & strSNP.VALUE.ToString & "','" & strPKGCODE.VALUE.ToString & "'," & _
                        " '" & txtCustomerCode.Text & "','" & gstrIpaddressWinSck & "','" & strSHIPPEDQTY.VALUE.ToString & "','" & strRECEIVEDQTY.VALUE.ToString & "','" & IIf(IsNothing(strTRIPNO.VALUE), "", strTRIPNO.VALUE).ToString & "')"

                    SqlConnectionclass.ExecuteNonQuery(strsql)
                    IsCompleted = True
                End If
            Next
            'shubhra
            strsql = "SELECT C.Cust_Drgno, COUNT(C.Item_code) " & _
                 " FROM CustItem_Mst C WHERE C.Account_CodE = '" & txtCustomerCode.Text & "'" & _
                 " AND C.UNIT_CODE = '" & gstrUNITID & "' AND C.Active = 1" & _
                 " AND EXISTS (SELECT TOP 1 C.Cust_Drgno FROM FORD_TRIGGER_FILE_NISSAN" & _
                 " WHERE UNIT_CODE = C.UNIT_CODE and CUST_CODE = C.Account_Code AND PartNo= C.Cust_Drgno )" & _
                 " GROUP BY C.Cust_Drgno HAVING COUNT(C.ITEM_CODE) > 1"
            oRdr = SqlConnectionclass.ExecuteReader(strsql)

            If oRdr.HasRows Then
                While oRdr.Read
                    strException = strException + oRdr("Cust_Drgno").ToString + " ,"
                End While
                If strException.Trim.Length > 0 Then
                    strException = Mid(strException, 1, strException.Length - 1)
                    MessageBox.Show("Following Customer Part No(s) are mapped with multiple Internal Part Codes:" + vbCrLf + strException, ResolveResString(100), MessageBoxButtons.OK)
                    IsCompleted = False
                    SqlConnectionclass.RollbackTran()
                    Exit Sub
                End If
            End If

            oRdr.Close() : oRdr = Nothing

            strsql = "UPDATE SYSFILE SET AUTORANSEQNO = " & intRANNo & "" & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "' and cust_code = '" & txtCustomerCode.Text & "'"

            SqlConnectionclass.ExecuteNonQuery(strsql)
            SqlConnectionclass.CommitTran()
            PB.Maximum = 100
            xlWorkBook.Close()

            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            strsql = "EXEC USP_AUTO_UPLOADFORDTRIGGER_NISSAN '" & gstrUNITID & "', '" & txtCustomerCode.Text & "', '" & gstrIpaddressWinSck & "'"
            SqlConnectionclass.ExecuteNonQuery(strsql)

            If IsCompleted = True Then
                MessageBox.Show("Trigger file Uploaded Successfully !", ResolveResString(100), MessageBoxButtons.OK)
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            PB.Maximum = 100
            PB.Visible = False
            KillExcelProcess(xlApp)
        End Try

    End Sub

    Private Function isOADate(ByVal Val As Object) As String

        Dim result As Double

        Try

            result = Convert.ToDouble(Val)
            Return "Y"

        Catch ex As Exception
            Return "N"
        End Try

    End Function

    Private Function isOAString(ByVal Val As Object) As String

        Dim result As String

        Try

            result = Convert.ToString(Val)
            Return "Y"

        Catch ex As Exception
            Return "N"
        End Try

    End Function

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
#End Region

    Private Sub cmdClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClear.Click

        Try
            txtCustomerCode.Text = ""
            lblCustName.Text = ""
            lblFileName.Text = ""

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    '10933955
    Private Sub optNissan_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optNissan.CheckedChanged
        Try

            If optNissan.Checked = True Then
                cmdCustHelp.Enabled = True
            Else
                txtCustomerCode.Enabled = False
                cmdCustHelp.Enabled = False
            End If

            cmdClear_Click(cmdClear, e)

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub cmdCustHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCustHelp.Click
        Dim strSQL As String = String.Empty
        Dim StrKeyValue As String = String.Empty
        Dim STRHELP() As String = Nothing
        '10933955
        Try

            If optNissan.Checked = True Then
                StrKeyValue = "Nissan"
            ElseIf optPCA.Checked = True Then
                StrKeyValue = "PCA"
            Else
                StrKeyValue = ""
            End If

            strSQL = "select DISTINCT L.Key2 CustomerCode, C.CUST_NAME CustomerName from Lists L Inner Join Customer_mst C " & _
                " ON L.UNIT_CODE = C.UNIT_CODE AND L.Key2 = C.Customer_Code " & _
                " where C.UNIT_CODE = '" & gstrUNITID & "' and L.Key1 = '" + StrKeyValue + "' "

            STRHELP = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Customer Help")

            If UBound(STRHELP) <> "-1" Then
                txtCustomerCode.Text = STRHELP(0)
                lblCustName.Text = STRHELP(1)
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub custHelpLblFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles custHelpLblFormat.Click
        '10933955
        Dim strSQL As String = String.Empty
        Dim STRHELP() As String = Nothing
        btnSave.Enabled = True
        Try
            strSQL = "select DISTINCT L.Key2 CustomerCode, C.CUST_NAME CustomerName from Lists L Inner Join Customer_mst C " & _
                " ON L.UNIT_CODE = C.UNIT_CODE AND L.Key2 = C.Customer_Code " & _
                " where C.UNIT_CODE = '" & gstrUNITID & "' and L.Key1 = 'Nissan' "

            STRHELP = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Customer Help")

            If UBound(STRHELP) <> "-1" Then
                txtCustLblFormat.Text = STRHELP(0)
                lblCustNameLblFormat.Text = STRHELP(1)
                PopulateItemsforMapping()
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating

        Dim strSQL As String = String.Empty
        Dim STRHELP() As String = Nothing
        '10933955
        Try
            strSQL = "select DISTINCT L.Key2 CustomerCode, C.CUST_NAME CustomerName from Lists L Inner Join Customer_mst C " & _
                " ON L.UNIT_CODE = C.UNIT_CODE AND L.Key2 = C.Customer_Code " & _
                " where C.UNIT_CODE = '" & gstrUNITID & "' and L.Key1 = 'Nissan' and  L.Key2 = '" & txtCustomerCode.Text & "'"

            STRHELP = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Customer Help")

            If UBound(STRHELP) <> "-1" Then
                txtCustomerCode.Text = STRHELP(0)
                lblCustName.Text = STRHELP(1)
            Else
                MessageBox.Show("Invalid Customer.", ResolveResString(100), MessageBoxButtons.OK)
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtCustLblFormat_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustLblFormat.Validating

        Dim strSQL As String = String.Empty
        Dim STRHELP() As String = Nothing
        Dim dt As New DataTable
        '10933955
        Try
            strSQL = "select DISTINCT L.Key2 CustomerCode, C.CUST_NAME CustomerName from Lists L Inner Join Customer_mst C " & _
                " ON L.UNIT_CODE = C.UNIT_CODE AND L.Key2 = C.Customer_Code " & _
                " where C.UNIT_CODE = '" & gstrUNITID & "' and L.Key1 = 'Nissan' and  L.Key2 = '" & Replace(txtCustLblFormat.Text, "'", "") & "'"
            dt = SqlConnectionclass.GetDataTable(strSQL)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                txtCustLblFormat.Text = Convert.ToString(dt.Rows(0)("CustomerCode"))
                lblCustNameLblFormat.Text = Convert.ToString(dt.Rows(0)("CustomerName"))
                'PopulateItemsforMapping()
            Else
                MessageBox.Show("Invalid Customer.", ResolveResString(100), MessageBoxButtons.OK)
                txtCustLblFormat.Text = ""
                sprMapping.MaxRows = 0
                lblCustNameLblFormat.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            dt.Dispose()
        End Try

    End Sub

    ' ''Private Sub SetSNPSpreadProperty()
    ' ''    'Added by MILIND--22/11/2017--101406664
    ' ''    Try
    ' ''        With sprItemSNP
    ' ''            .DisplayRowHeaders = True
    ' ''            .MaxRows = 0
    ' ''            .MaxCols = 0
    ' ''            .MaxCols = enm_SNPmapping.SNPFormat
    ' ''            .UserColAction = FPSpreadADO.UserColActionConstants.UserColActionSort
    ' ''            .Row = 0
    ' ''            .Col = enm_SNPmapping.DrgNo : .Text = "Part Number" : .set_ColWidth(enm_mapping.DrgNo, 20)
    ' ''            .Col = enm_SNPmapping.Desc : .Text = "Part Description" : .set_ColWidth(enm_mapping.Desc, 35)
    ' ''            .Col = enm_SNPmapping.SNPFormat : .Text = "Label Format" : .set_ColWidth(enm_mapping.Format, 25)
    ' ''            .set_RowHeight(.Row, 15)
    ' ''        End With
    ' ''    Catch ex As Exception
    ' ''        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    ' ''    Finally
    ' ''        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    ' ''    End Try
    ' ''End Sub
    Private Sub SetSpreadProperty()
        Try
            With sprMapping
                .DisplayRowHeaders = True
                .MaxRows = 0
                .MaxCols = 0
                .MaxCols = enm_mapping.SupplyArea
                .UserColAction = FPSpreadADO.UserColActionConstants.UserColActionSort
                .Row = 0
                .Col = enm_mapping.DrgNo : .Text = "Part Number" : .set_ColWidth(enm_mapping.DrgNo, 2500)
                .Col = enm_mapping.Desc : .Text = "Part Description" : .set_ColWidth(enm_mapping.Desc, 4000)
                .Col = enm_mapping.Format : .Text = "Label Size" : .set_ColWidth(enm_mapping.Format, 2200)
                .Col = enm_mapping.SNPFormat : .Text = "Label Type" : .set_ColWidth(enm_mapping.SNPFormat, 1700)
                .Col = enm_mapping.ReceivingArea : .Text = "Receiving Area" : .set_ColWidth(enm_mapping.SNPFormat, 1700)
                .Col = enm_mapping.SupplyArea : .Text = "Supply Area" : .set_ColWidth(enm_mapping.SNPFormat, 1700)
                .set_RowHeight(.Row, 400)
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    '  --101421656-
    ''Private Sub SetSNPSpreadColType(ByVal pintRowNo As Integer)

    ' ''    'Added by MILIND--22/11/2017--101406664
    ' ''    Try
    ' ''        With sprItemSNP
    ' ''            .Row = pintRowNo
    ' ''            .Col = enm_SNPmapping.DrgNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
    ' ''            .Col = enm_SNPmapping.Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
    ' ''            .Col = enm_SNPmapping.SNPFormat : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
    ' ''        End With
    ' ''    Catch ex As Exception
    ' ''        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    ' ''    Finally
    ' ''        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    ' ''    End Try
    ' ''End Sub
    'Private Sub SetSNPSpreadColTypes(ByVal pintRowNo As Integer)
    '    Try
    '        With sprItemSNP
    '            .Row = pintRowNo
    '            .Col = enm_SNPmapping.DrgNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
    '            .Col = enm_SNPmapping.Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
    '            .Col = enm_SNPmapping.SNPFormat : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
    '        End With
    '    Catch ex As Exception
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    Finally
    '        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Sub

    'End Here--MILIND--101406664
    Private Sub SetSpreadColTypes(ByVal pintRowNo As Integer)
        Try
            With sprMapping
                .Row = pintRowNo
                .Col = enm_mapping.DrgNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enm_mapping.Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enm_mapping.Format : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enm_mapping.SNPFormat : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    'Private Sub addRowAtEnterKeyPressSNP()         --101421656-
    '    'Added by MILIND --22/11/2017--101406664
    '    Dim oRDR As SqlDataReader = Nothing
    '    Dim strSQL As String = Nothing
    '    '10933955
    '    Try

    '        With sprItemSNP
    '            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .set_RowHeight(.Row, 15)
    '            .Col = enm_SNPmapping.DrgNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '            .Col = enm_SNPmapping.Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '            strSQL = "select format from Lists where UNIT_CODE = '" & gstrUNITID & "'" & _
    '                " AND Key1 = 'NissanLabels' AND Key2 = '" & txtCustLblFormat.Text & "'"
    '            oRDR = SqlConnectionclass.ExecuteReader(strSQL)

    '            If oRDR.HasRows Then
    '                .TypeComboBoxList = " " & Chr(9)
    '                While oRDR.Read
    '                    .Row = .MaxRows : .Col = enm_SNPmapping.SNPFormat
    '                    .TypeComboBoxList = .TypeComboBoxList & oRDR("format").ToString & Chr(9)
    '                End While
    '            End If

    '        End With
    '    Catch ex As Exception
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    Finally
    '        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Sub
    ' ''Private Sub addSNPRowAtEnterKeyPress()
    ' ''    Dim oRDR As SqlDataReader = Nothing
    ' ''    Dim strSQL As String = Nothing
    ' ''    'Added by MILIND--22/11/2017--101406664
    ' ''    Try
    ' ''        With sprItemSNP
    ' ''            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
    ' ''            .MaxRows = .MaxRows + 1
    ' ''            .Row = .MaxRows
    ' ''            .set_RowHeight(.Row, 18)
    ' ''            Call SetSNPSpreadColType(.Row)

    ' ''            strSQL = "select descr from Lists where UNIT_CODE = '" & gstrUNITID & "'" & _
    ' ''                " AND Key1 = 'isLabelCountSNP' AND Key2 = '" & txtCustomerCodeSNP.Text & "'"
    ' ''            oRDR = SqlConnectionclass.ExecuteReader(strSQL)

    ' ''            If oRDR.HasRows Then
    ' ''                .TypeComboBoxList = " " & Chr(9)
    ' ''                While oRDR.Read
    ' ''                    .Row = .MaxRows : .Col = enm_SNPmapping.SNPFormat
    ' ''                    .TypeComboBoxList = .TypeComboBoxList & oRDR("descr").ToString & Chr(9)
    ' ''                End While
    ' ''            End If

    ' ''        End With
    ' ''    Catch ex As Exception
    ' ''        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    ' ''    Finally
    ' ''        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    ' ''    End Try
    ' ''End Sub
    Private Sub addRowAtEnterKeyPress()
        Dim oRDR As SqlDataReader = Nothing
        Dim strSQL As String = Nothing
        '10933955
        Try
            With sprMapping
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 350)
                Call SetSpreadColTypes(.Row)

                strSQL = "select descr from Lists where UNIT_CODE = '" & gstrUNITID & "'" & _
                    " AND Key1 = 'NissanLabels' AND Key2 = '" & txtCustLblFormat.Text & "'"


                oRDR = SqlConnectionclass.ExecuteReader(strSQL)

                If oRDR.HasRows Then
                    .Row = .MaxRows : .Col = enm_mapping.Format
                    .TypeComboBoxList = " " & Chr(9)
                    While oRDR.Read
                        .Row = .MaxRows : .Col = enm_mapping.Format
                        .TypeComboBoxList = .TypeComboBoxList & oRDR("descr").ToString & Chr(9)
                    End While
                End If

                ' --101421656-
                'Call SetSpreadColTypes(.Row)
                '.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                '.MaxRows = .MaxRows + 1
                '.Row = .MaxRows
                '.set_RowHeight(.Row, 350)
                'Call SetSpreadColTypes(.Row)

                strSQL = "select descr from Lists where UNIT_CODE = '" & gstrUNITID & "'" & _
                 " AND Key1 = 'isLabelCountSNP' AND Key2 = '" & txtCustLblFormat.Text & "'"
                oRDR = Nothing
                oRDR = SqlConnectionclass.ExecuteReader(strSQL)

                If oRDR.HasRows Then
                    .Row = .MaxRows : .Col = enm_mapping.SNPFormat
                    .TypeComboBoxList = " " & Chr(9)
                    While oRDR.Read
                        .Row = .MaxRows : .Col = enm_mapping.SNPFormat
                        .TypeComboBoxList = .TypeComboBoxList & oRDR("descr").ToString & Chr(9)
                    End While
                End If
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    ' ''Private Sub TabPage3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage3.Click
    ' ''    'Added by MILIND--22/11/2017--101406664
    ' ''    Try
    ' ''        SetSNPSpreadProperty()

    ' ''    Catch ex As Exception
    ' ''        RaiseException(ex)
    ' ''    End Try
    ' ''End Sub 

    ' --101421656- 
    Private Sub TabPage2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Click
        Try
            If txtCustLblFormat.Text = "" Then
                SetSpreadProperty()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    '  --101421656-
    ''Private Sub PopulateItemsforSNPMapping()
    ' ''    'Added by MILIND--22/11/2017--101406664
    ' ''    Dim strSQL As String = ""
    ' ''    Dim oRdr As SqlDataReader = Nothing
    ' ''    Dim objSNPformat As Object = Nothing

    ' ''    Try
    ' ''        sprItemSNP.MaxRows = 0
    ' ''        strSQL = "SELECT top 1 * " & _
    ' ''            " FROM Lists l inner join CustItem_Mst c ON L.UNIT_CODE = C.UNIT_CODE " & _
    ' ''            " AND L.Key2 = C.Account_Code WHERE L.UNIT_CODE = '" & gstrUNITID & "'" & _
    ' ''            " AND L.Key2 = '" & txtCustomerCodeSNP.Text & "' AND Active = 1"

    ' ''        If IsRecordExists(strSQL) Then
    ' ''            strSQL = "SELECT Distinct C.CUST_DRGNO, C.DRG_DESC,c.isLabelCountSNP" & _
    ' ''            " FROM Lists l inner join CustItem_Mst c ON L.UNIT_CODE = C.UNIT_CODE " & _
    ' ''            " AND L.Key2 = C.Account_Code WHERE L.UNIT_CODE = '" & gstrUNITID & "'" & _
    ' ''            " AND L.Key2 = '" & txtCustomerCodeSNP.Text & "' AND Active = 1"

    ' ''            oRdr = SqlConnectionclass.ExecuteReader(strSQL)
    ' ''            If oRdr.HasRows Then
    ' ''                While oRdr.Read
    ' ''                    addSNPRowAtEnterKeyPress()
    ' ''                    With sprItemSNP
    ' ''                        .Row = .MaxRows
    ' ''                        .Col = enm_SNPmapping.DrgNo : .Text = oRdr("CUST_DRGNO").ToString
    ' ''                        .Col = enm_SNPmapping.Desc : .Text = oRdr("DRG_DESC").ToString
    ' ''                        objSNPformat = oRdr("isLabelCountSNP").ToString
    ' ''                        If objSNPformat = True Then
    ' ''                            .Col = enm_SNPmapping.SNPFormat : .Text = "SNP"
    ' ''                        Else
    ' ''                            .Col = enm_SNPmapping.SNPFormat : .Text = "Normal"
    ' ''                        End If
    ' ''                    End With
    ' ''                End While
    ' ''            End If
    ' ''        End If
    ' ''    Catch ex As Exception
    ' ''        RaiseException(ex)
    ' ''    End Try
    ' ''End Sub   '
    '--101421656-
    Private Sub PopulateItemsforMapping()
        Dim strSQL As String = ""
        Dim oRdr As SqlDataReader = Nothing
        Dim objSNPformat As Object = Nothing

        Try
            sprMapping.MaxRows = 0
            strSQL = "SELECT top 1 * " & _
                " FROM Lists l inner join CustItem_Mst c ON L.UNIT_CODE = C.UNIT_CODE " & _
                " AND L.Key2 = C.Account_Code WHERE L.UNIT_CODE = '" & gstrUNITID & "'" & _
                " AND L.Key2 = '" & txtCustLblFormat.Text & "' AND Active = 1"

            If IsRecordExists(strSQL) Then
                strSQL = "SELECT Distinct C.CUST_DRGNO, C.DRG_DESC,c.labelFormat Format,c.isLabelCountSNP,c.RECVAREANISSAN,c.SUPPLYAREANISSAN" & _
                " FROM Lists l inner join CustItem_Mst c ON L.UNIT_CODE = C.UNIT_CODE " & _
                " AND L.Key2 = C.Account_Code WHERE L.UNIT_CODE = '" & gstrUNITID & "'" & _
                " AND L.Key2 = '" & txtCustLblFormat.Text & "' AND Active = 1"

                oRdr = SqlConnectionclass.ExecuteReader(strSQL)
                If oRdr.HasRows Then
                    While oRdr.Read
                        addRowAtEnterKeyPress()
                        With sprMapping
                            .Row = .MaxRows
                            .Col = enm_mapping.DrgNo : .Text = oRdr("CUST_DRGNO").ToString
                            .Col = enm_mapping.Desc : .Text = oRdr("DRG_DESC").ToString
                            .Col = enm_mapping.Format : .Text = oRdr("Format").ToString
                            objSNPformat = oRdr("isLabelCountSNP").ToString
                            .Col = enm_mapping.ReceivingArea : .Text = oRdr("RECVAREANISSAN").ToString
                            .Col = enm_mapping.SupplyArea : .Text = oRdr("SUPPLYAREANISSAN").ToString
                            If objSNPformat = True Then
                                .Col = enm_mapping.SNPFormat : .Text = "SNP"
                            Else
                                .Col = enm_mapping.SNPFormat : .Text = "Normal"
                            End If
                        End With
                    End While
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim strSql As String = Nothing
        Dim objDrgNo As Object = Nothing
        Dim objFormat As Object = Nothing
        Dim objFormatSNP As Object = Nothing
        Dim intRow As Integer = 0
        Dim objChk As Object = Nothing
        Dim objRecvAr As Object = Nothing
        Dim objSuppAr As Object = Nothing
        Try
            With sprMapping
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = enm_mapping.DrgNo : objDrgNo = .Text
                    .Col = enm_mapping.Format : objFormat = .Text
                    .Col = enm_mapping.SNPFormat : objFormatSNP = .Text
                    .Col = enm_mapping.ReceivingArea : objRecvAr = .Text
                    .Col = enm_mapping.SupplyArea : objSuppAr = .Text
                    strSql = "Update Custitem_Mst set labelFormat = '" & objFormat & "',IsLabelCountSNP=" & IIf(Convert.ToString(objFormatSNP) = "SNP", 1, 0) & " , RecvAreaNissan = '" & objRecvAr & "', SupplyAreaNissan = '" & objSuppAr & "'" & _
                       " where unit_code = '" & gstrUNITID & "'" & _
                       " and account_code = '" & txtCustLblFormat.Text & "'" & _
                       " and cust_drgno = '" & objDrgNo & "'"

                    SqlConnectionclass.ExecuteNonQuery(strSql)
                Next
            End With
            If strSql = "" Then
                MsgBox("Please select customer Code !", MsgBoxStyle.Information)
            Else
                'MessageBox.Show("Label Formats Updated", ResolveResString(100), MsgBoxStyle.Information, MessageBoxButtons.OK)
                MsgBox("Label Formats Updated !", MsgBoxStyle.Information)
            End If

            PopulateItemsforMapping()

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ' --101421656-
    ''Private Sub cmdItemSNP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdItemSNP.Click
    ' ''    'Added by MIlIND--22/11/2017--101406664
    ' ''    Dim strSQL As String = String.Empty
    ' ''    Dim STRHELP() As String = Nothing

    ' ''    Try
    ' ''        strSQL = "select DISTINCT L.Key2 CustomerCode, C.CUST_NAME CustomerName from Lists L Inner Join Customer_mst C " & _
    ' ''            " ON L.UNIT_CODE = C.UNIT_CODE AND L.Key2 = C.Customer_Code " & _
    ' ''            " where C.UNIT_CODE = '" & gstrUNITID & "' and L.Key1 = 'Nissan' "

    ' ''        STRHELP = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Customer Help")

    ' ''        If UBound(STRHELP) <> "-1" Then
    ' ''            txtCustomerCodeSNP.Text = STRHELP(0)
    ' ''            lblCustmerCodeSNP.Text = STRHELP(1)
    ' ''            PopulateItemsforSNPMapping()
    ' ''        End If

    ' ''    Catch ex As Exception
    ' ''        RaiseException(ex)
    ' ''    End Try
    ' ''End Sub


    ' ''Private Sub btnSNPSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSNPSave.Click
    ' ''    'Added by MILIND--22/11/2017--101406664
    ' ''    Dim strSql As String = Nothing
    ' ''    Dim objDrgNo As Object = Nothing
    ' ''    Dim objFormat As Object = Nothing
    ' ''    Dim objFormatBit As Object = Nothing
    ' ''    Dim intRow As Integer = 0
    ' ''    Dim objChk As Object = Nothing

    ' ''    Try
    ' ''        With sprItemSNP
    ' ''            For intRow = 1 To .MaxRows
    ' ''                .Row = intRow
    ' ''                .Col = enm_SNPmapping.DrgNo : objDrgNo = .Text
    ' ''                .Col = enm_SNPmapping.SNPFormat : objFormat = .Text
    ' ''                If objFormat = "SNP" Then
    ' ''                    objFormatBit = 1
    ' ''                Else
    ' ''                    objFormatBit = 0
    ' ''                End If
    ' ''                strSql = "Update Custitem_Mst set isLabelCountSNP  ='" & objFormatBit & "' " & _
    ' ''                   " where unit_code = '" & gstrUNITID & "'" & _
    ' ''                   " and account_code = '" & txtCustomerCodeSNP.Text & "'" & _
    ' ''                   " and cust_drgno = '" & objDrgNo & "'"

    ' ''                SqlConnectionclass.ExecuteNonQuery(strSql)
    ' ''            Next
    ' ''        End With
    ' ''        MessageBox.Show("Label Formats Updated", ResolveResString(100), MessageBoxButtons.OK)
    ' ''        PopulateItemsforSNPMapping()

    ' ''    Catch ex As Exception
    ' ''        RaiseException(ex)
    ' ''    End Try
    ' 
    ''End Sub
    ' --101421656-

    Private Sub optPCA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPCA.CheckedChanged
        Try

            If optPCA.Checked = True Then
                cmdCustHelp.Enabled = True
            Else
                txtCustomerCode.Enabled = False
                cmdCustHelp.Enabled = False
            End If

            cmdClear_Click(cmdClear, e)

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub GETDataFromTextFile_PCA(ByVal filepath As String)
        'Data part for DataTable
        Dim dtPCAData As New DataTable("Table")
        If dtPCAData.Columns.Count = 0 Then
            dtPCAData.Columns.Add("SenderID", GetType(String))              '0
            dtPCAData.Columns.Add("ReceiverID", GetType(String))            '1
            dtPCAData.Columns.Add("MessageCallOffNo", GetType(String))      '2
            dtPCAData.Columns.Add("MessageCallOffDate", GetType(String))    '3
            dtPCAData.Columns.Add("ConsigneeID", GetType(String))           '4
            dtPCAData.Columns.Add("SellerID", GetType(String))              '5
            dtPCAData.Columns.Add("ConsignorID", GetType(String))           '6
            dtPCAData.Columns.Add("AccountID", GetType(String))             '7
            dtPCAData.Columns.Add("BuyersPartNumber", GetType(String))      '8
            dtPCAData.Columns.Add("PlacePortCode", GetType(String))         '9
            dtPCAData.Columns.Add("ReferenceNo", GetType(String))           '10
            dtPCAData.Columns.Add("Quantity", GetType(String))              '11
            dtPCAData.Columns.Add("Numbers", GetType(String))               '12
            dtPCAData.Columns.Add("ScheduleDeliveryDate", GetType(String))  '13
            dtPCAData.Columns.Add("PackageCode", GetType(String))           '14
            dtPCAData.Columns.Add("placeofdestination", GetType(String))    '15
            dtPCAData.Columns.Add("uploadingpoint", GetType(String))    '16
        End If

        Dim arr(16) As String
        'Ending of Data Part
        Dim dtUploadDataStatus As New DataTable("T")
        Dim tfp As New TextFieldParser(filepath)
        Dim fields As String()

        Dim StrSql As String = ""

        Try
            tfp.Delimiters = New String() {","}
            tfp.TextFieldType = FieldType.Delimited

            'tfp.ReadLine() ' skip header
            While tfp.EndOfData = False
                fields = tfp.ReadFields()

                'Checking For Invalid file
                If fields.Length <> 17 Then
                    MessageBox.Show("Please upload a valid file!", "Invalid File", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    lblFileName.Text = ""
                    Exit Sub
                End If
                'Ending of Checking For Invalid file

                'Checking Valid file on 02_Aug_2022- Client Shared this file dont have header
                'Flling Data Table from CSV file data.
                arr(0) = fields(0)  'SenderID
                arr(1) = fields(1)  'ReceiverID
                arr(2) = fields(2)  'MessageCallOffNo
                arr(3) = fields(3)  'MessageCallOffDate
                arr(4) = fields(4)  'ConsigneeID
                arr(5) = fields(5)  'SellerID
                arr(6) = fields(6)  'ConsignorID
                arr(7) = fields(7)  'AccountID
                arr(8) = fields(8)  'BuyersPartNumber
                arr(9) = fields(9)  'PlacePortCode
                arr(10) = fields(10) 'ReferenceNo
                arr(11) = fields(11) 'Quantity
                arr(12) = fields(12) 'Numbers
                arr(13) = fields(13) 'ScheduleDeliveryDate
                arr(14) = fields(14) 'PackageCode
                arr(15) = fields(15) 'placeofdestination
                arr(16) = fields(16) 'UPLOADINGPOINT

                dtPCAData.Rows.Add(arr)
            End While

            If dtPCAData IsNot Nothing AndAlso dtPCAData.Rows.Count > 0 Then
                Dim sqlCmd As New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 300 ' 5 Minute
                    .CommandText = "USP_PCA_SCHEDULE_UPLOADING"
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@PCAData", dtPCAData)
                    .Parameters.AddWithValue("@CustomerCode", txtCustomerCode.Text)
                    .Parameters.AddWithValue("@UnitCode", gstrUNITID)
                    .Parameters.AddWithValue("@Ip_Address", gstrIpaddressWinSck)
                    .Parameters.AddWithValue("@UserId", mP_User.ToString.Trim)
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    dtUploadDataStatus = SqlConnectionclass.GetDataTable(sqlCmd)
                    If Convert.ToString(.Parameters("@MSG").Value) <> "" Then
                        MsgBox(Convert.ToString(Replace(.Parameters("@MSG").Value, "\n", Environment.NewLine)), MsgBoxStyle.Exclamation, ResolveResString(100))
                    Else
                        MsgBox("Data Uploaded Successfully.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    End If
                End With
            Else
                MsgBox("File Does not have Data.", MsgBoxStyle.Exclamation, ResolveResString(100))
            End If
        Catch ex As Exception
            MessageBox.Show("Please upload a valid file!", "Invalid File", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblFileName.Text = ""
        Finally
            lblFileName.Text = ""
            tfp.Dispose()
        End Try

    End Sub
End Class