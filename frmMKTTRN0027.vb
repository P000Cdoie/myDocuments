Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0027
    Inherits System.Windows.Forms.Form
    Dim gblintCount As Integer = 0
	'*****************************************************************************
	'Copyright©2002             - MIND
	'Form Name (Physical Name)  - frmMKTTRN0027.frm
	'Created by                 - Nitin Sood
	'Created Date               - 13-01-2003
	'Revision Date              - 12/02/04
	'Revision History           - Schedule can not move in loaded folder.
	'Revision Date              - 20/09/2004
	'Revision History           - New Function for Entry in Daily Marketing Schedule by Sourabh Khatri.
	'Form Description           - Upload Schedules from MSSL into Database
	' Changes done by Sourabh on 02 Sep 2004 For Add two new column DSNO and DSDATETIME(DSTracking-10623)
	' Changes done by sourabh on 23 sep 2004 for add new procedure for changing the functionality of daily schedule uploading
	' Changes done by sourabh on 25 sep 2004 correction in Despatch update Query
	' Changes done by Nisha on 27 sep 2004 correction in Despatch update Query
	' Changes done by sourabh on 28 sep 2004 correction in Despatch update Query
	'Revision Date              - 29/11/2006
	'Revision History           - Save Weekday of schedule and current date as Ent_date.
	'                           - Issue Id:19114 for DS wise Planned / Unplanned schedule status report.
	'Revision Date      :         08 Feb 2008
	'Revised By         :         Prashant Rajpal
	'Revision for       :        Issue No.22228- Schedule uploading
    'Modified by Amit on 05/May/2011 for multiunit change
	'*****************************************************************************
    'revision by :prashant Rajpal
    'revision on 10th and 11 th dec 2012
    'issue id  10318214 
    ' smiel multi, unit migration changes
    '*****************************************************************************
    'Revision Date      :    30 jan 2013
    'Revised By         :    Prashant Rajpal
    'Revision for       :    Issue No.10331478
    'purpose            :   EDI Functionality for Mate Unit3  export     
    '*****************************************************************************
    'Revision Date      :    15-mar-2013- 18-mar-2013
    'Revised By         :    Prashant Rajpal
    'Revision for       :    10354980 
    'purpose            :    Woco Migration changes
    '***********************************************************************************
    'Revision Date      :    04-Feb-2015
    'Revised By         :    Prashant Rajpal
    'Revision for       :    10754868 
    'purpose            :    MATE BANGALORE CHANGES 
    '***********************************************************************************

	Private mlngFormTag As Short 'Form Tag
	Private Enum StatusTabTypes
		sScheduleFolder = 0
		sCheckStatus = 1
	End Enum
	Private udtItemSchedule() As MKTSchedules 'ARRAY UDT To Store Item Code and Schedule Date
	Private Sub cmdclose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdclose.Click
		On Error GoTo ErrHandler
		Me.Close()
		Exit Sub
ErrHandler: 
		gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mp_connection)
	End Sub
	
	Private Sub cmdTransfer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTransfer.Click
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		'Call Buttons Event to Transfer Schedules
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		On Error GoTo ErrHandler
		Dim intCount As Short 'For Next Loop ...
		Dim intTotCount As Short
		
		intTotCount = 0
        For intCount = 0 To lstStatus.Items.Count - 1
            If lstStatus.Items.Item(intCount).Checked = True Then
                intTotCount = intTotCount + 1
            End If
        Next
        If intTotCount = 1 Then
            'changed by : prashant rajpal as on 26 may 2008
            Me.lstStatus.Enabled = False
            cmdTransfer.Enabled = False
            cmdClose.Enabled = False
            Call cmdUploadSchd_ButtonClick(cmdUploadSchd, New UCActXCtl.UCfraRepCmd.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT))
        Else
            MsgBox("Atleast Check One Schedule/Maximum one Schedule Can Check to be Updated in Empower.", MsgBoxStyle.Information, "empower")
            If lstStatus.Enabled = True Then
                lstStatus.Focus()
            Else
                cmdClose.Focus()
            End If
        End If

        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        'changed by : prashant rajpal as on 26 may 2008
        Me.lstStatus.Enabled = True
    End Sub
    Private Sub TransferCheckedSchedules()
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
        'Created By     -   Nitin Sood
        'Description    -   Transfer Files that are Checked in List
        '*-*-*-*-*-*-*-*-*-*-*--*-*-*-*-*-*-*-*-*-*-*-*-*-*-**-*-*-*-*-*-
        On Error GoTo ErrHandler
        Dim Intcounter As Short 'For Next Loop . . .
        Dim strReturn As String
        Dim strErrorMessage As String
        For Intcounter = 0 To lstStatus.Items.Count - 1
            If lstStatus.Items.Item(Intcounter).Checked Then
                'Only For Checked Files
                strReturn = TransferFile(Trim(lstStatus.Items.Item(Intcounter).Text))
                If Mid(strReturn, 1, 1) = "Y" Then
                    'File Transfered Successfully
                    lstStatus.Items.Item(Intcounter).Font = VB6.FontChangeBold(lstStatus.Items.Item(Intcounter).Font, True)
                    lstStatus.Items.Item(Intcounter).ForeColor = System.Drawing.Color.Red
                    MsgBox(" Schedule Uploaded Successfully .")
                    If cmdTransfer.Enabled = False Then
                        cmdClose.Enabled = True
                        cmdTransfer.Enabled = True
                    End If
                    'changed by : prashant rajpal as on 26 may 2008
                    Me.lstStatus.Enabled = True
                Else
                    If Len(strReturn) > 2 Then
                        strErrorMessage = "File Could not be Uploaded in empower due to following reason(s) :- " & vbCrLf & Mid(strReturn, InStr(1, strReturn, "»") + 1, Len(strReturn))
                        MsgBox(strErrorMessage, MsgBoxStyle.Information, "empower")
                        strErrorMessage = ""
                        If cmdTransfer.Enabled = False Then
                            cmdClose.Enabled = True
                            cmdTransfer.Enabled = True
                        End If
                        cmdClose.Focus()
                        'changed by : prashant rajpal as on 26 may 2008
                        Me.lstStatus.Enabled = True
                    End If
                End If
            End If
        Next
        'changed by : prashant rajpal as on 26 may 2008
        Me.lstStatus.Enabled = True
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub GetStartDateEndDate(ByVal strGiveFileName As String, ByRef StrSTdate As Object, ByRef strEndDt As String)
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
        'Created By     -   Nitin Sood
        'Description    -   Returns Back the Start Date and End Date of Schedule Files iin the passed Parameters
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
        On Error GoTo ErrHandler
        System.Windows.Forms.Application.DoEvents()
        StrSTdate = Mid(strGiveFileName, 4, 4) ' 4 Chars (ddmm)
        strEndDt = Mid(strGiveFileName, 8, 4) ' 4 Chars (ddmm)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
   
    Private Function TransferFile(ByVal strFileName As String) As String
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
        'Created By     -   Nitin Sood
        'Description    -   Transfer Individual File
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
        On Error GoTo ErrHandler
        Dim strStartDate As String
        Dim strEndDate As String
        Dim strResult As String
        Call GetStartDateEndDate(strFileName, strStartDate, strEndDate)
        'Change Mouse Pointer
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        'Check whether Monthly or Daily Schedule
        Select Case LCase(VB.Right(strFileName, 4))
            Case ".sch" 'It is Daily Schedule.
                'Upload Daily Schedule
                'Call UpdateDAILYSchedule(strFileName)
                strResult = DailyScheduleEntry(strFileName)
            Case ".msh" 'It is Monthly Schedule.
                'Upload MOnthly scheduel
                If MsgBox("This is a Monthly Schedule file [ " & strFileName & " ]. Wish to continue...?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "empower") = MsgBoxResult.Yes Then
                    strResult = UpdateMonthSchedule(strFileName)
                End If
            Case Else
                strResult = strResult & "Invalid file name or file is in wrong format [" & strFileName & "]."
        End Select
        TransferFile = strResult
        'Change Mouse Pointer to default
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        cmdClose.Enabled = True
        cmdTransfer.Enabled = True
        Exit Function
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

Private Function UpdateMonthSchedule(ByVal strMonthlySchdFileName As String) As String
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*-*-*
        'Created By     -   Nitin Sood
        'Description    -   Upload Monthly Schedule In ForeCast Master
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*-*-*
        On Error GoTo ErrHandler
        Dim strCurDtSQL As String 'Todays Date
        Dim strFileDtTm As String 'Creation Date And Time Of File being Uploaded
        Dim nInSchFile As Short 'Free File Number
        Dim intRows As Short
        Dim intRowNo As Short
        Dim counter, intPos As Short
        Dim strsql As String
        Dim strMaster As String
        Dim strSchPath, strRow, strBuffer As String
        Dim strMstFields() As String
        Dim strFields(5) As String
        Dim stracccode As String 'Customer Code
        Dim strschno As String
        Dim strYYMM As String 'Date in YYYYMM Format
        Dim strStartYYYYmm As String 'Date in YYYYMM Format Start Schedule
        Dim strEndYYYYmm As String
        Dim blnUpdfg As Boolean 'Boolean For Again Upload the Same File
        Dim strItemCode As String
        Dim intRowSeperator As Short
        Dim dblDispatchqty As Double
        Dim nInJunkFile As Short 'To Open Junk File
        Dim ass2 As Scripting.FileSystemObject
        Dim ass3 As Scripting.FileSystemObject
        Dim a As Scripting.TextStream
        Dim xmlBasic As Scripting.TextStream
        Dim intFileNumber As Short
        Dim strmail As String
        Dim strTemp As String
        Dim ooutlook As Object
        Dim blnjunkgenerated As Boolean
        Dim strErrorMessage As String
        Dim intErrorNumber As Short
        Dim strString As String
        Dim intRevisionNo As Short
        Dim sch_itemcode As String
        Dim inttotal_custitems As Short
        ass2 = New Scripting.FileSystemObject
        ass3 = New Scripting.FileSystemObject 'XML
        blnjunkgenerated = False
        System.Windows.Forms.Application.DoEvents()
        If gobjDB.GetResult("SELECT CONVERT(VARCHAR(10),GETDATE(),103)+ SPACE(1) +LEFT(CONVERT(VARCHAR(10),GETDATE(),108),5) AS Dt") = False Then GoTo ErrHandler
        strCurDtSQL = gobjDB.GetValue("dt")
        'prashant rajpal change ended
        strSchPath = txtschdPath.Text & "\" '.SCH file Path
        'Check Date & Time of File Creation, assume it as customer update time
        strFileDtTm = VB6.Format(FileDateTime(strSchPath & strMonthlySchdFileName), "dd/MMM/yyyy hh:mm:ss")
        FileClose(nInSchFile)
        nInSchFile = FreeFile() 'Get Free File Number
        FileOpen(nInSchFile, strSchPath & strMonthlySchdFileName, OpenMode.Input)
        System.Windows.Forms.Application.DoEvents() 'Pass Control
        '------------------------
        'a = ass2.OpenTextFile(strSchPath & "JunkFile.jnk", Scripting.IOMode.ForWriting, True)
        'xmlBasic = ass3.OpenTextFile(strSchPath & "JunkFile.xml", Scripting.IOMode.ForWriting, True) 'XML
        'xmlBasic.WriteLine(("<?xml version='1.0' encoding='ISO-8859-1'?>")) 'XML
        'xmlBasic.WriteLine(("<Junk>")) 'XMl
        ''------------------------
        'Read the file and store all the values in strBuffer
        strBuffer = LineInput(nInSchFile)
        If InStr(strBuffer, "##") > 0 Then
            strMaster = Mid(strBuffer, 1, InStr(strBuffer, "##") - 1) 'Meta Data for File
            strBuffer = Mid(strBuffer, InStr(strBuffer, "##") + 2)
        Else
            intErrorNumber = intErrorNumber + 1 'Add In error String
            strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid file, Master Data not found." & vbCrLf
            UpdateMonthSchedule = "N»" & strErrorMessage
            Exit Function
        End If
        System.Windows.Forms.Application.DoEvents()
        intRowNo = 0 : intRows = 0
        intRows = UBound(Split(strBuffer, "^"))
        ' Master Data
        counter = 0 : intPos = 0
        While Len(strMaster) > 0
            counter = counter + 1
            intPos = InStr(strMaster, "|")
            ReDim Preserve strMstFields(counter)
            If intPos > 0 Then
                strMstFields(counter) = Mid(strMaster, 1, intPos - 1)
            Else
                strMstFields(counter) = strMaster
                strMaster = ""
            End If
            strMaster = Mid(strMaster, intPos + 1)
        End While
        System.Windows.Forms.Application.DoEvents()
        'Check Vendor Code At MSSL (Fix) S073 is of SMIEL
        If gstrUNITID = "SML" Then
            If UCase(strMstFields(1)) <> "S073" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If
        If gstrUNITID = "SMR" Then
            If UCase(strMstFields(1)) <> "S073T" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        '10754868
        If gstrUNITID = "MST" Then
            If UCase(strMstFields(1)) <> "M581" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        '10754868
        If gstrUNITID = "M03" Then
            If UCase(strMstFields(1)) <> "M582" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        If gstrUNITID = "MAE" Then
            If UCase(strMstFields(1)) <> "M554" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        If gstrUNITID = "MAN" Then
            If UCase(strMstFields(1)) <> "M1117" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        If gstrUNITID = "MEN" Then
            If UCase(strMstFields(1)) <> "M1117G" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        If gstrUNITID = "SED" Then
            If UCase(strMstFields(1)) <> "S073E" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strMonthlySchdFileName & "]. Please Check vendor code." & vbCrLf
                UpdateMonthSchedule = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        'Check MSSL Unit Code,And Get Account Code From Customer Master based on ScheduleCode
        strMstFields(4) = Trim(UCase(strMstFields(4)))

        stracccode = SelectDataFromTable("Customer_Code", "Customer_Mst", " ScheduleCode = '" & strMstFields(4) & "'")
        If Trim(stracccode) = "" Then
            intErrorNumber = intErrorNumber + 1
            strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid unit code " & strMstFields(4) & " in File  [" & strMonthlySchdFileName & "]." & vbCrLf
            UpdateMonthSchedule = "N»" & strErrorMessage
            If cmdTransfer.Enabled = False Then
                cmdClose.Enabled = True
                cmdTransfer.Enabled = True
            End If
            Exit Function
        End If

        'Junk Data
        'a.Write(("Junk Data Received From Customer - " & stracccode & " in FILE - " & strMonthlySchdFileName))
        'xmlBasic.WriteLine(("<empower>Junk Data Received From Customer - " & stracccode & " in FILE - " & strMonthlySchdFileName & "</empower>"))
        'Convert date format
        'strMstFields(3) = VB6.Format(strMstFields(3), gstrDateFormat) 'Schedule Made On
        'strMstFields(5) = VB6.Format(strMstFields(5), gstrDateFormat) 'Schedule Start Date
        'strMstFields(6) = VB6.Format(strMstFields(6), gstrDateFormat) 'Schedule Enddate Date
        strMstFields(3) = VB6.Format(strMstFields(3), "dd-MMM-yyyy") 'Schedule Made On
        strMstFields(5) = VB6.Format(strMstFields(5), "dd-MMM-yyyy") 'Schedule Start Date
        strMstFields(6) = VB6.Format(strMstFields(6), "dd-MMM-yyyy") 'Schedule Enddate Date
        strschno = strMstFields(2) ' set the value of Schedule NO
        strYYMM = VB6.Format(strCurDtSQL, "yyyymm")
        strStartYYYYmm = VB6.Format(strMstFields(5), "YYYYMM")
        strEndYYYYmm = VB6.Format(strMstFields(6), "YYYYMM")
        'If CDate(strStartYYYYmm) > Format(GetServerDate(), "YYYYMM") Then
        If VB6.Format(CDate(strMstFields(5)), "yyyymm") < VB6.Format(GetServerDate(), "YYYYMM") Then
            MsgBox("Schedule can not be uploaded for previous month ", MsgBoxStyle.OkOnly, ResolveResString(100))
            'a.Close()
            'xmlBasic.Close()
            FileClose(nInSchFile)
            If cmdTransfer.Enabled = False Then
                cmdClose.Enabled = True
                cmdTransfer.Enabled = True
            End If
            Exit Function
        End If
        System.Windows.Forms.Application.DoEvents()
        'Code add by sourabh for make control number on client dependent
        If UBound(strMstFields) < 7 Then
            ReDim Preserve strMstFields(8)
            strMstFields(7) = CStr(0)
            strMstFields(8) = CStr(0)
        End If
        If IsNumeric(strMstFields(8)) = False Then
            UpdateMonthSchedule = "N»Control Number Must Be Numeric Value "
            If cmdTransfer.Enabled = False Then
                cmdClose.Enabled = True
                cmdTransfer.Enabled = True
            End If
            Exit Function
        End If
        strString = ValidateControlNo(strschno, stracccode, CInt(strMstFields(8)))
        If VB.Left(strString, 1) = "N" Then
            UpdateMonthSchedule = strString
            If cmdTransfer.Enabled = False Then
                cmdClose.Enabled = True
                cmdTransfer.Enabled = True
            End If
            Exit Function
        End If
        'Check whether file already uploaded          HAS TO BE DONE
        blnUpdfg = False ' True if Updating, false if new
        If ValidateData("SchCustUpdDt", "Sched_Smiel_Mst", " Customer_Code = '" & stracccode & "' AND ScheduleNo = '" & strschno & "' AND ScheduleDate = '" & strMstFields(3) & "' AND SchStartDate = '" & strMstFields(5) & "' AND SchEndDate = '" & strMstFields(6) & "' and UNIT_code='" & gstrUNITID & "'") Then
            System.Windows.Forms.Application.DoEvents()
            If MsgBox("File [ " & strMonthlySchdFileName & " ] is already uploaded." & Chr(13) & "Wish to upload again...?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "empower") <> MsgBoxResult.Yes Then
                If cmdTransfer.Enabled = False Then
                    cmdClose.Enabled = True
                    cmdTransfer.Enabled = True
                End If
                Exit Function
            End If
            blnUpdfg = True 'UPLOAD AGAIN
        End If
        'Begin Trans
        'mP_Connection.BeginTrans
        mP_Connection.Execute("Set dateFormat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If blnUpdfg = True Then

            'Delete Details from MonthlyMKTSchedule And RevisionNo = 0 ,
            '   Year_Month Between strStartYYYYMM and strEndYYYYMM
            strsql = "Delete From MonthlyMKTSchedule Where Account_Code = '" & stracccode & "' And RevisionNo = 0 And Year_Month Between '" & strStartYYYYmm & "' AND '" & strEndYYYYmm & "' and unit_code='" & gstrUNITID & "'"
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            'Update Revision No -1 for Rest of Schedules for the Same Account in Same Range
            strsql = " UPDATE MonthlyMKTSchedule Set RevisionNo = RevisionNo - 1 Where Account_Code = '" & stracccode & "' And Year_Month Between '" & strStartYYYYmm & "' AND '" & strEndYYYYmm & "' and unit_code='" & gstrUNITID & "'"
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            strsql = "Insert Into Sched_Smiel_Mst (Customer_Code, ScheduleNo, ScheduleDate, SchStartDate, SchEndDate, SchUpdDate, "
            strsql = strsql & " SchCustUpdDt, SchEntryDate,Control_No,Unit_Code ) Values ( '" & stracccode & "', '"
            strsql = strsql & strMstFields(2) & "', '" & strMstFields(3) & "', '"
            strsql = strsql & strMstFields(5) & "', '" & strMstFields(6) & "', "
            strsql = strsql & " '" & strCurDtSQL & "' ,  '" & strFileDtTm & "','" & strCurDtSQL & "'," & strMstFields(8) & ",'" + gstrUNITID + "')"
        Else
            'Update Master Table
            strsql = "Insert Into Sched_Smiel_Mst (Customer_Code, ScheduleNo, ScheduleDate, SchStartDate, SchEndDate, SchUpdDate, "
            strsql = strsql & " SchCustUpdDt, SchEntryDate,Control_No,Unit_Code ) Values ( '" & stracccode & "', '"
            strsql = strsql & strMstFields(2) & "', '" & strMstFields(3) & "', '"
            strsql = strsql & strMstFields(5) & "', '" & strMstFields(6) & "', "
            strsql = strsql & " '" & strCurDtSQL & "' ,  '" & strFileDtTm & "','" & strCurDtSQL & "'," & strMstFields(8) & ",'" + gstrUNITID + "')"
        End If
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        'strsql = "Delete from forecast_mst_LatestHistory where  unit_code='" & gstrUNITID & "'"
        'mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        Dim Intcounter As Short
        Intcounter = 0
        strsql = " Select * From  forecast_mst Where Unit_Code='" + gstrUNITID + "' And customer_code =  '" & stracccode & "' and due_date >= '" & strMstFields(5) & "'  and   Enagare_UNLOC= 'N/A' "
        gobjDB.GetResult(strsql)
        inttotal_custitems = gobjDB.RowCount
        mP_Connection.Execute("Set dateFormat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        gobjDB.MoveFirst()
        'for update revision no  ie revision no=revision no +1
        mP_Connection.Execute("UPDATE Forecast_Mst_History Set RevisionNo = RevisionNo + 1 Where  Customer_Code = '" & stracccode & "'  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        While Not gobjDB.EOFRecord
            mP_Connection.Execute("insert into forecast_mst_history (Customer_code,Product_no,Due_date,Quantity,RevisionNo,ScheduleNo,ent_dt,ent_userid,upd_dt,upd_userid,enagare_unloc,unit_Code ) " & " Values ('" & stracccode & "' ,'" & gobjDB.GetValue("product_no") & "', '" & gobjDB.GetValue("due_date") & "', '" & gobjDB.GetValue("Quantity") & "', '0'," & " '" & gobjDB.GetValue("scheduleno") & "','" & gobjDB.GetValue("ent_dt") & "', '" & gobjDB.GetValue("ent_userid") & "', '" & gobjDB.GetValue("upd_dt") & "' ,'" & gobjDB.GetValue("upd_userid") & "','N/A','" + gstrUNITID + "')   ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            gobjDB.MoveNext()
        End While

        mP_Connection.Execute("delete From  forecast_mst Where customer_code =  '" & stracccode & "' and due_date >= '" & strMstFields(5) & "'  and   Enagare_UNLOC= 'N/A'  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        While Len(strBuffer) > 0 'Till END of Remaining Buffer
            intRowSeperator = InStr(strBuffer, "^") 'Done By Nitin Sood   8 May 2003
            System.Windows.Forms.Application.DoEvents()
            counter = 0 : intRowNo = intRowNo + 1
            strRow = Mid(strBuffer, 1, intRowSeperator - 1)
            While Len(strRow) > 0
                counter = counter + 1 : intPos = InStr(strRow, "|")
                If intPos > 0 Then
                    strFields(counter) = Mid(strRow, 1, intPos - 1)
                Else
                    strFields(counter) = strRow
                    strRow = ""
                End If
                strRow = Mid(strRow, intPos + 1)
            End While
            'Convert date format
            strYYMM = VB6.Format(strFields(2), "yyyymm") 'yyyymm Get Month Year
            'strFields(2) = VB6.Format(strFields(2), gstrDateFormat) 'From date
            'strFields(3) = VB6.Format(strFields(3), gstrDateFormat) 'To Date (THIS DATE IS TO BE USED FOR STORAGE IN DUE DATE FIELD IN FORECAST MASTER)
            strFields(2) = VB6.Format(strFields(2), "dd-MMM-yyyy") 'From date
            strFields(3) = VB6.Format(strFields(3), "dd-MMM-yyyy") 'To Date (THIS DATE IS TO BE USED FOR STORAGE IN DUE DATE FIELD IN FORECAST MASTER)
            'CustFrwg Number is Supplied , and Stored in strFields(1) Array
            ' Get Item Code
            If Not ValidateData("CD.Cust_DrgNo", " Cust_Ord_Dtl CD,Cust_Ord_Hdr CH ", " CD.unit_Code = CH.unit_Code and CD.Account_Code = CH.Account_Code AND CD.Cust_Ref = CH.Cust_Ref AND CD.Amendment_No = CH.Amendment_No AND  CD.Account_Code = '" & stracccode & "' AND CD.Cust_DrgNo = '" & strFields(1) & "' AND CD.ACTIVE_Flag = 'A' AND CH.PO_Type = 'O' and cd.UNIT_code='" & gstrUNITID & "'") Then
                'Customer Drwg No Not Found
                MsgBox("Customer Drawing Number " & strFields(1) & " not found." & vbCr & "Adding this Drawing No. in JunkFile." & vbCr & "Continuing with rest of the Schedule File...", MsgBoxStyle.Information, "empower")
                'a.WriteLine((" Customer Drawing No. - " & strFields(1)))
                'xmlBasic.WriteLine(("<Cust_DrgNo>" & strFields(1) & "</Cust_DrgNo>"))
                blnjunkgenerated = True
                GoTo ItemCodeNotFound
            Else
                'Get Internal Item Code
                strItemCode = Trim(SelectDataFromTableNew("Item_Code", " Cust_Ord_Dtl CD,Cust_Ord_Hdr CH ", " CD.unit_Code = CH.unit_Code and  CD.Account_Code = CH.Account_Code AND CD.Cust_Ref = CH.Cust_Ref AND CD.Amendment_No = CH.Amendment_No AND  CD.Account_Code = '" & stracccode & "' AND CD.Cust_DrgNo = '" & strFields(1) & "' AND CD.ACTIVE_Flag = 'A' AND CH.PO_Type = 'O' and cd.UNIT_code='" & gstrUNITID & "'"))
            End If
            System.Windows.Forms.Application.DoEvents()
            'Check whether Item already exists,Monthly Data will always Exist in Forecast IT MAY OR MAY NOT EXIST IN monthlyMKTSchedule
            If (ValidateData("Product_No", "Forecast_Mst", " Customer_Code = '" & stracccode & "' AND Due_Date = '" & strFields(3) & "' And Product_No = '" & strItemCode & "' and UNIT_code='" & gstrUNITID & "'")) Or blnUpdfg = True Then
                'If Entry for this Item is already there .... for strFields(3) Date
                'Check in Forecast_Mst_History ,If Entry Exists there Also
                If (ValidateData("Product_No", "Forecast_Mst_History", " Customer_Code = '" & stracccode & "' AND Due_Date = '" & strFields(3) & "' And Product_No = '" & strItemCode & "' and UNIT_code='" & gstrUNITID & "'")) Or blnUpdfg = True Then
                End If
                'Insert Item into Forecast Mst
                strsql = "Insert Into ForeCast_Mst (Customer_code,product_no,Due_date,Quantity,ent_dt,ent_userid"
                strsql = strsql & ",upd_dt,upd_userid ,scheduleno,Unit_Code) Values ( '" & stracccode & "', '" & strItemCode & "', "
                strsql = strsql & "'" & Trim(strFields(3)) & "', "
                strsql = strsql & Val(strFields(4))
                strsql = strsql & ",'" & strCurDtSQL & "' ,'PORT','" & strCurDtSQL & "','PORT','" & strschno & "','" + gstrUNITID + "')"
            Else
                'Check in Forecast_Mst_History ,If Entry Exists there
                If ValidateData("Product_No", "Forecast_Mst_History", " Customer_Code = '" & stracccode & "' AND Due_Date = '" & strFields(3) & "' And Product_No = '" & strItemCode & "' and UNIT_code='" & gstrUNITID & "'") Then
                    'Update All Previous Histories Here with Revision No = Revision No + 1
                    '1.)Update and INCREASE by 1 Revision No for Previous Entries
                    strsql = "UPDATE Forecast_Mst_History Set RevisionNo = RevisionNo + 1 Where "
                    strsql = strsql & " Customer_Code = '" & stracccode & "' AND Due_Date = '" & strFields(3) & "' AND Product_No = '" & strItemCode & "' and unit_code='" & gstrUNITID & "'"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'UPDATE REVISION NOs
                Else
                    '            strSql = ""
                    '    strsql = "Insert Into ForeCast_Mst_LatestHistory (Customer_code,product_no,Due_date,Quantity,RevisionNo,ScheduleNo,ent_dt,ent_userid"
                    '    strsql = strsql & ",upd_dt,upd_userid ,Unit_Code) Values ( '" & stracccode & "', '" & strItemCode & "', "
                    '    strsql = strsql & "'" & Trim(strFields(3)) & "', "
                    '    strsql = strsql & Val(strFields(4))
                    '    strsql = strsql & ",0,'" & strschno & "' ,'" & strCurDtSQL & "' ,'PORT','" & strCurDtSQL & "','PORT','" + gstrUNITID + "' )"
                    '    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'UPDATE REVISION NOs
                End If
                strsql = "Insert Into ForeCast_Mst_History (Customer_code,Product_no,Due_date,Quantity,RevisionNo,ScheduleNo,ent_dt,ent_userid,upd_dt,upd_userid,enagare_unloc,Unit_code ) "
                strsql = strsql & " Select '" & stracccode & "', '" & strItemCode & "','" & Trim(strFields(3)) & "'," & Val(strFields(4))
                strsql = strsql & ",0,scheduleno,'" & strCurDtSQL & "' ,'PORT','" & strCurDtSQL & "','PORT','N/A',Unit_Code "
                strsql = strsql & " from forecast_mst where Customer_Code = '" & stracccode & "' and Product_no = '" & strItemCode & "' AND Due_Date = '" & strFields(3) & "' and unit_code='" & gstrUNITID & "'"
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                'Insert Item into Forecast Mst
                strsql = "Insert Into ForeCast_Mst (Customer_code,product_no,Due_date,Quantity,ent_dt,ent_userid"
                strsql = strsql & ",upd_dt,upd_userid ,scheduleno,Unit_Code) Values ( '" & stracccode & "', '" & strItemCode & "', "
                strsql = strsql & "'" & Trim(strFields(3)) & " ', "
                strsql = strsql & Val(strFields(4))
                strsql = strsql & ",'" & strCurDtSQL & "' ,'PORT','" & strCurDtSQL & "','PORT','" & strschno & "','" + gstrUNITID + "' )"
            End If
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'Forecast Mst SQL
ItemCodeNotFound:
            If InStr(strBuffer, "^") > 0 Then
                strBuffer = Mid(strBuffer, InStr(strBuffer, "^") + 1)
            End If
            If Len(strBuffer) = 0 Or strBuffer = "^" Then strBuffer = ""
        End While
        strsql = "Delete from forecast_mst_historyNew where unit_code='" & gstrUNITID & "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strsql = "Insert into  forecast_mst_historyNew"
        strsql = strsql & " (Customer_code,product_no,Due_date,Quantity,ScheduleNo,RevisionNo,ent_userid,ent_dt,upd_userid,upd_dt,Unit_Code)"
        strsql = strsql & " Select Customer_code,product_no,Due_date,Quantity,'" & strMstFields(2) & "' as scheduleNo,0 as RevisionNo,ent_userid,ent_dt,upd_userid,upd_dt,Unit_Code"
        strsql = strsql & " From forecast_mst Where Not Exists"
        strsql = strsql & " (Select Customer_code,product_no,Due_date, Quantity,ScheduleNo,"
        strsql = strsql & "  RevisionNo, ent_userid,ent_dt, upd_userid,"
        strsql = strsql & "  upd_dt from forecast_mst_history Where Unit_Code='" + gstrUNITID + "' And customer_code = forecast_mst.customer_code"
        strsql = strsql & " and Due_Date = forecast_mst.Due_date"
        strsql = strsql & " and product_no = forecast_mst.Product_no"
        strsql = strsql & " and Due_Date between '" & strMstFields(5) & "' and '" & strMstFields(6) & "'"
        strsql = strsql & " and Customer_Code = '" & stracccode & "' and ScheduleNo = '" & strMstFields(2) & "'"
        strsql = strsql & " Union"
        strsql = strsql & " Select Customer_code,product_no,Due_date, Quantity,ScheduleNo,RevisionNo, ent_userid,ent_dt, upd_userid,upd_dt from forecast_mst_Latesthistory Where Unit_Code='" + gstrUNITID + "' And customer_code = forecast_mst.customer_code"
        strsql = strsql & " and Due_Date = forecast_mst.Due_date"
        strsql = strsql & " and product_no = forecast_mst.Product_no"
        strsql = strsql & " and Due_Date between '" & strMstFields(5) & "' and '" & strMstFields(6) & "'"
        strsql = strsql & " and Customer_Code = '" & stracccode & "' and ScheduleNo = '" & strMstFields(2) & "'"
        strsql = strsql & " )and Due_Date between '" & strMstFields(5) & "' and '" & strMstFields(6) & "' and Customer_Code = '" & stracccode & "' and unit_code='" & gstrUNITID & "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        System.Windows.Forms.Application.DoEvents()
        FileClose(nInSchFile) 'Close File
        'Sending mail using MAPI
        'Code add by Sourabh
        '---------------------------------------
        'xmlBasic.WriteLine(("</Junk>"))
        'xmlBasic.Close() 'Close File
        'a.Close()
        'xmlBasic.Close
        '---------------------------------------
        'mP_Connection.CommitTrans
        'Check Whether a file with the same Name already exists
        If Len(Dir(strSchPath & "Loaded Files\" & strMonthlySchdFileName)) > 0 Then
            'Remove existing file
            Kill(strSchPath & "Loaded Files\" & strMonthlySchdFileName)
        End If
        Dim ass As Scripting.FileSystemObject
        ass = New Scripting.FileSystemObject
        'Move file to Folder "\Loaded Files"
        If ass.FileExists(strSchPath & strMonthlySchdFileName) Then
            ass.MoveFile(strSchPath & strMonthlySchdFileName, strSchPath & "Loaded Files\" & strMonthlySchdFileName)
        End If
        'If blnjunkgenerated = True Then
        'ass.MoveFile(strSchPath & "JunkFile.jnk", strSchPath & "Loaded Files\" & VB.Left(strMonthlySchdFileName, Len(strMonthlySchdFileName) - 4) & "JunkFile.jnk")
        'End If
        UpdateMonthSchedule = "Y»"
        ass = Nothing
        Exit Function
ErrHandler:
        If Err.Number = 53 Then
            MsgBox(" File not found ", MsgBoxStyle.OkOnly, ResolveResString(100))
            If cmdTransfer.Enabled = False Then
                cmdClose.Enabled = True
                cmdTransfer.Enabled = True
            End If
        End If
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function  
    Private Sub dirSchdFiles_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dirSchdFiles.Change
        On Error GoTo ErrHandler
        txtschdPath.Text = dirSchdFiles.Path
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub drvSchds_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles drvSchds.SelectedIndexChanged
        On Error GoTo ErrHandler
        dirSchdFiles.Path = drvSchds.Drive 'Pick Drive from DriveSchedule
        Exit Sub
ErrHandler:
        If Err.Number = 68 Then
            MsgBox("Drive Selected is NOT available.Choose another Drive.", MsgBoxStyle.Information, "empower")
            drvSchds.Focus() : Exit Sub
        Else
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub frmMKTTRN0027_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
       
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mlngFormTag
        frmModules.NodeFontBold(Me.Tag) = True
        If lstStatus.Enabled = False Then
            cmdClose.Focus() 'As All Ctrls are Disabled
        Else
            lstStatus.Focus()
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0027_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '-----------------------------------------------------------------------------
        'Created By     -   Nitin Sood
        'Description    -   Makes Font in ListView in frmModule as NORMAL.
        '-----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0027_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '-----------------------------------------------------------------------
        'Escape Key Handling
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be unloaded
                If MsgBox("Want To Close This Screen ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "empower") = MsgBoxResult.Yes Then
                    Me.Close()
                    GoTo EventExitSub
                Else
                    Me.ActiveControl.Focus()
                    GoTo EventExitSub
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0027_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------------
        'Created By     -   Nitin Sood
        'Description    -   Call HELP if F4 key is pressed
        '-----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        ' On pressing F4 , help gets dispayed
        If (KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0) Then Call ctlUploadSchedulesHDR_Click(ctlUploadSchedulesHDR, New System.EventArgs())
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0027_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.AppStarting)
        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlUploadSchedulesHDR.Tag)
        Call FitToClient(Me, fraMain, ctlUploadSchedulesHDR, lblUploadCmd, 300)
        'Setting Print and Close Buttons
        cmdTransfer.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lblUploadCmd.Left) + 70)
        cmdTransfer.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblUploadCmd.Top) + 50)
        cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdTransfer.Left) + VB6.PixelsToTwipsX(cmdTransfer.Width) + 10)
        cmdClose.Top = cmdTransfer.Top
        Call InitializeFormSettings() 'Initial Form Settings
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub InitializeFormSettings()
        On Error GoTo ErrHandler
        Dim Intcounter As Short
        Dim intCounterdir As Short
        Dim strDefaultPath As String
        With lstStatus
            .View = System.Windows.Forms.View.Details
            .GridLines = True
            .Enabled = False
            .FullRowSelect = True
            .HotTracking = True
            .HoverSelection = True
            .CheckBoxes = True
            .MultiSelect = False
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            .Columns.Add("         Schedule Files")
            .Columns.Item(0).Width = VB6.TwipsToPixelsX(7480)
        End With
        ssUploadSchd.TabPages.Item(StatusTabTypes.sScheduleFolder).Enabled = False 'Disable ScheduleFolder TAB
        ssUploadSchd.SelectedIndex = StatusTabTypes.sCheckStatus
        Select Case ShowNewSchedules()
            Case 10
                MsgBox("Vendor Schedule Folders not found in Application path.Contact System Administrator.", MsgBoxStyle.Information, "empower")
        End Select
        ReDim udtItemSchedule(0) 'Initialize Array
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Function ShowNewSchedules() As Short
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
        'Checks For New Schedule at The Location PreDefined
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
        On Error GoTo ErrHandler
        Dim strSchPath As String
        Dim strbackpath As String
        Dim strBackpathRowWise() As String
        Dim strNextFile As String
        ShowNewSchedules = 0 'No Problems
        Me.cmdTransfer.Enabled = False
        txtschdPath.Text = Trim(Find_Value("SELECT ISNULL(MSSL_INBOUND_PATH,'') FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'"))
        'txtschdPath.Text = "G:\Vendor_schedule\sed_path\"
        If gstrUNITID = "SML" Then
            'txtschdPath.Text = My.Application.Info.DirectoryPath '"
            strbackpath = txtschdPath.Text
            strbackpath = Replace(strbackpath, "\\", "\")
            strBackpathRowWise = Split(strbackpath, "\")
            strbackpath = "\\" & strBackpathRowWise(1) & "\Vendor_Schedule"
        End If

        If gstrUNITID = "SMR" Then
            strbackpath = txtschdPath.Text
            strBackpathRowWise = Split(strbackpath, "\")
        End If

        If gstrUNITID = "M03" Then
            strbackpath = txtschdPath.Text
            strBackpathRowWise = Split(strbackpath, "\")
        End If
        If gstrUNITID = "MAE" Then
            strbackpath = txtschdPath.Text
            strBackpathRowWise = Split(strbackpath, "\")
        End If
        '10754868
        If gstrUNITID.ToUpper = "MST" Then
            strbackpath = txtschdPath.Text
            strBackpathRowWise = Split(strbackpath, "\")
        End If
        If gstrUNITID = "MAN" Then
            strbackpath = txtschdPath.Text
            strBackpathRowWise = Split(strbackpath, "\")
        End If

        If gstrUNITID = "MEN" Then
            strbackpath = txtschdPath.Text
            strBackpathRowWise = Split(strbackpath, "\")
        End If
        If gstrUNITID = "SED" Then
            'txtschdPath.Text = My.Application.Info.DirectoryPath '"
            strbackpath = txtschdPath.Text
            strbackpath = Replace(strbackpath, "\\", "\")
            strBackpathRowWise = Split(strbackpath, "\")
            strbackpath = "\\" & strBackpathRowWise(1) & "\Vendor_Schedule\\Exception Files"
        End If
        '10754868
        txtschdPath.Text = strbackpath
        'strSchPath = Dir(txtschdPath.Text, FileAttribute.Directory)
        strSchPath = strbackpath
        strNextFile = Dir(txtschdPath.Text & "\", FileAttribute.Directory)
        If Trim(strSchPath) <> "" Then
            'Some Files are found
            While strNextFile <> ""
                System.Windows.Forms.Application.DoEvents()
                'Reject . and .. directories
                'If FileName$ Like ".SCH" Then
                '23 july 2018 
                'If InStr(strNextFile, ".sch") <> 0 Or InStr(strNextFile, ".msh") <> 0 Then
                'lstStatus.Items.Add(strNextFile) 'Add to List
                'End If
                If gstrUNITID = "MAN" Then
                    If InStr(strNextFile, ".sch") <> 0 Then
                        lstStatus.Items.Add(strNextFile) 'Add to List
                    End If
                ElseIf gstrUNITID = "SED" Then
                    If Mid(strNextFile, 12, 5) = "S073E" And (InStr(strNextFile, ".sch") <> 0 Or InStr(strNextFile, ".msh") <> 0) Then
                        lstStatus.Items.Add(strNextFile) 'Add to List
                    End If
                Else
                    If InStr(strNextFile, ".sch") <> 0 Or InStr(strNextFile, ".msh") <> 0 Then
                        lstStatus.Items.Add(strNextFile) 'Add to List
                    End If
                    End If
                    '23 july 2018
                    strNextFile = Dir()
            End While
            If lstStatus.Items.Count = 0 Then 'No Schedules were found
                MsgBox("No Requirement file exists.", MsgBoxStyle.Information, "empower")
                With lstStatus
                    .Enabled = False
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                End With
                Me.cmdTransfer.Enabled = False 'Disable ListStatus and Transfer Button
            Else
                With lstStatus
                    .Enabled = True
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End With
                Me.cmdTransfer.Enabled = True
            End If
        Else
            'Path was not Correct.
            ShowNewSchedules = 10
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub frmMKTTRN0027_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        'Assign form to nothing
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ssUploadSchd_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ssUploadSchd.SelectedIndexChanged
        Static PreviousTab As Short = ssUploadSchd.SelectedIndex()
        On Error GoTo ErrHandler
       
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        PreviousTab = ssUploadSchd.SelectedIndex()
    End Sub
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '------------------------------------------------------------------------------
        'Created By     -   Nitin Sood
        'Description    -   Get Data from BackEnd
        '------------------------------------------------------------------------------
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & " Where Unit_Code='" + gstrUNITID + "' And " & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTable = ""
            End If
        Else
            SelectDataFromTable = ""
        End If
        GetDataFromTable.ResultSetClose()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function SelectDataFromTableNew(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '------------------------------------------------------------------------------
        'Created By     -   Nitin Sood
        'Description    -   Get Data from BackEnd
        '------------------------------------------------------------------------------
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & " Where " & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                SelectDataFromTableNew = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTableNew = ""
            End If
        Else
            SelectDataFromTableNew = ""
        End If
        GetDataFromTable.ResultSetClose()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function ValidateData(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCond As String) As Boolean
        '-----------------------------------------------------------------------------------------
        'Created By     -   Nitin Sood
        'Description    -   Validates Data
        'ReturnType     -   True    - If data is Correct
        '               -   False   - If Data is not Correct
        '-----------------------------------------------------------------------------------------
        Dim strsql As String
        Dim clsInstValidate As New ClsResultSetDB
        On Error GoTo ErrHandler
        strsql = "Select " & mstrFieldName & " From " & mstrTableName & " Where  " & mstrCond
        If clsInstValidate.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If clsInstValidate.GetNoRows > 0 Then
                ValidateData = True 'If Data Found
            Else
                ValidateData = False 'If data Not Found
            End If
        Else
            ValidateData = False 'If data Not Found
        End If
        clsInstValidate.ResultSetClose() 'Close Recordset
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function DailyScheduleEntry(ByRef strSchFILE As String) As String
        On Error GoTo Errorhandler
        '-----------------------------------------------------------------------------------------
        'Created By     -   Sourabh Khatri
        'Description    -   Data Uploading in DailymktSchedule
        'ReturnType     -   True    - If data uploading is Correct
        '               -   False   - If Data is not uploaded Correctly
        '-----------------------------------------------------------------------------------------
        Dim strRow, strMaster, strBuffer As String
        Dim strMstFields() As String
        Dim strFields(7) As String
        Dim strSbuItCode As String
        Dim intRowSeperator, intPos, counter, nInSchFile As Short
        Dim strFileDtTm As String 'To Save File Dat Time
        Dim intRows, intRowNo As Short ' For Looping
        Dim blnUpdfg As Boolean ' Flag True if Schedule Already Uploaded.
        Dim strCurDtSQL As Date
        Dim stracccode As String 'Customer Code
        Dim strschno As String 'Schedule No
        Dim strItemCode As String 'Item Code
        Dim Intcounter As Short 'For Next Loop ...
        Dim lnRowCounter As Integer 'For Count How many Items in Text File
        Dim lnInsertCounter As Integer 'For Count how many items inserted in Database
        Dim ass2 As Scripting.FileSystemObject 'Object for XML File
        Dim ass3 As Scripting.FileSystemObject 'Object for XML File
        Dim intFileNumber As Short 'integer variable for free file no
        Dim strmail As String 'For mail address
        Dim strTemp As String 'for temp mail address
        Dim ooutlook As Object 'Outlook Application object
        Dim blnJunkgenerate As Boolean
        Dim strSchPath As String
        Dim a As Scripting.TextStream
        Dim strErrorMessage As String
        Dim intErrorNumber As Short
        Dim xmlBasic As Scripting.TextStream
        Dim strDSNo As String
        Dim dtDatetime As Date
        Dim rsDSTracking As New ClsResultSetDB
        Dim blnDSTracking As Boolean
        Dim rsUpload As New ADODB.Recordset
        Dim strString As String
        Dim strControlValidate As String
        Dim strsql As String
        Dim intYYYYMM As Integer
        Dim strUpdateSchedule As String
        Dim strDeleteScheduleHistory As String
        Dim strSched_Smiel_mst As String
        Dim strDailymktSchedule As String
        Dim strDailySchedule_History As String
        Dim strDailymktUpdate As String
        Dim strMonthlymktSchedule As String
        ass2 = New Scripting.FileSystemObject
        ass3 = New Scripting.FileSystemObject 'XML
        blnJunkgenerate = False
        System.Windows.Forms.Application.DoEvents()
        blnJunkgenerate = False
        'strCurDtSQL = CDate(VB6.Format(GetServerDate, gstrDateFormat))
        strCurDtSQL = getDateForDB(GetServerDate)
        strSchPath = txtschdPath.Text & "\" '.SCH file Path
        FileClose(nInSchFile)
        nInSchFile = FreeFile() 'Get Free File Number
        FileOpen(nInSchFile, strSchPath & strSchFILE, OpenMode.Input, , OpenShare.Shared)
        System.Windows.Forms.Application.DoEvents() 'Pass Control
        '------------------------
        a = ass2.OpenTextFile(strSchPath & "JunkFile.jnk", Scripting.IOMode.ForWriting, True)
        xmlBasic = ass3.OpenTextFile(strSchPath & "JunkFile.xml", Scripting.IOMode.ForWriting, True) 'XML
        xmlBasic.WriteLine(("<?xml version='1.0' encoding='ISO-8859-1'?>")) 'XML
        xmlBasic.WriteLine(("<Junk>")) 'XMl
        '------------------------
        'Read the file and store all the values in strBuffer and Master
        strBuffer = LineInput(nInSchFile)
        If InStr(strBuffer, "##") > 0 Then
            strMaster = Mid(strBuffer, 1, InStr(strBuffer, "##") - 1) 'Meta Data for File
            strBuffer = Mid(strBuffer, InStr(strBuffer, "##") + 2)
        Else
            intErrorNumber = intErrorNumber + 1 'Add In error String
            strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid file, Master Data not found." & vbCrLf
            DailyScheduleEntry = "N»" & strErrorMessage
            Exit Function
        End If
        FileClose(nInSchFile)
        '------ Check Whether DailySchedule Amendment
        intRowSeperator = InStr(strBuffer, "^")
        strRow = Mid(strBuffer, 1, intRowSeperator - 1)
        intRows = UBound(Split(strRow, "|"))
        If intRows <> 6 Then '6 Array Elements SHOULD be made
            If intRows <> 5 Then
                intErrorNumber = intErrorNumber + 1 'Add To error List
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid file structure[" & strSchFILE & "].File couldn't be transferred." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
            'Stop DailySchedule Amendments temporarily
            intErrorNumber = intErrorNumber + 1
            strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Daily Schedule Amendment is not allowed for File [" & strSchFILE & "]." & vbCrLf
            DailyScheduleEntry = "N»" & strErrorMessage
            Exit Function
        End If
        System.Windows.Forms.Application.DoEvents()
        intRowNo = 0 : intRows = 0 'Re-Initialize
        lnRowCounter = UBound(Split(strBuffer, "^")) 'No.8 Of Entries
        lnInsertCounter = 0
        ' Master Data
        counter = 0 : intPos = 0
        While Len(strMaster) > 0
            counter = counter + 1
            intPos = InStr(strMaster, "|")
            ReDim Preserve strMstFields(counter)
            If intPos > 0 Then
                strMstFields(counter) = Mid(strMaster, 1, intPos - 1) 'Master Data Array
            Else
                strMstFields(counter) = strMaster
                strMaster = ""
            End If
            strMaster = Mid(strMaster, intPos + 1)
        End While
        System.Windows.Forms.Application.DoEvents()
        'Check Vendor Code At MSSL (Fix) S073 is of SMIEL
        If gstrUNITID = "SML" Then
            If UCase(strMstFields(1)) <> "S073" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If
        If gstrUNITID = "SMR" Then
            If UCase(strMstFields(1)) <> "S073T" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        If gstrUNITID = "MST" Then
            If UCase(strMstFields(1)) <> "M581" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        If gstrUNITID = "M03" Then
            If UCase(strMstFields(1)) <> "M582" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If
        If gstrUNITID = "MAE" Then
            If UCase(strMstFields(1)) <> "M554" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If
        If gstrUNITID = "MAN" Then
            If UCase(strMstFields(1)) <> "M1117" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If
        If gstrUNITID = "MEN" Then
            If UCase(strMstFields(1)) <> "M1117G" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If


        If gstrUNITID = "SED" Then
            If UCase(strMstFields(1)) <> "S073E" Then
                intErrorNumber = intErrorNumber + 1
                strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid File [" & strSchFILE & "]. Please Check vendor code." & vbCrLf
                DailyScheduleEntry = "N»" & strErrorMessage
                Exit Function
            End If
        End If

        'Check MSSL Unit Code,And Get Account Code From Customer Master based on ScheduleCode
        strMstFields(4) = UCase(strMstFields(4))
        stracccode = SelectDataFromTable("Customer_Code", "Customer_Mst", " ScheduleCode = '" & strMstFields(4) & "'")
        If Trim(stracccode) = "" Then
            intErrorNumber = intErrorNumber + 1
            strErrorMessage = strErrorMessage & CStr(intErrorNumber) & ").Invalid unit code " & strMstFields(4) & " in File  [" & strSchFILE & "]." & vbCrLf
            DailyScheduleEntry = "N»" & strErrorMessage
            Exit Function
        End If
        'Junk Data
        a.Write(("Junk Data Received From Customer - " & stracccode & " in FILE - " & strSchFILE))
        xmlBasic.WriteLine(("<empower>Junk Data Received From Customer - " & stracccode & " in FILE - " & strSchFILE & "</empower>"))
        System.Windows.Forms.Application.DoEvents()
        'Convert date format
        strMstFields(3) = Mid(strMstFields(3), 1, 23) 'Schedule Made On
        strMstFields(5) = Mid(strMstFields(5), 1, 10) 'Schedule Start Date Date
        strMstFields(6) = Mid(strMstFields(6), 1, 10) 'Schedule End Date Date
        strschno = Trim(strMstFields(2)) ' Schedule NO
        strFileDtTm = strMstFields(3)
        System.Windows.Forms.Application.DoEvents()
        'Code add by sourabh for make control number on client dependent
        If UBound(strMstFields) < 7 Then
            ReDim Preserve strMstFields(8)
            strMstFields(7) = CStr(0)
            strMstFields(8) = CStr(0)
        End If
        If IsNumeric(CInt(strMstFields(8))) = False Then
            DailyScheduleEntry = "N»Control Number Must Be Numeric Value "
            Exit Function
        End If
        strControlValidate = ValidateControlNo(strschno, stracccode, CInt(strMstFields(8)))
        If VB.Left(strControlValidate, 1) = "N" Then
            DailyScheduleEntry = strControlValidate
            Exit Function
        End If
        'Validation for Re Upload Schedule
        blnUpdfg = False
        If ValidateData("ScheduleNo", "Sched_Smiel_Mst", " Customer_Code = '" & stracccode & "' and ScheduleNo = '" & strschno & "' and UNIT_code='" & gstrUNITID & "'") Then
            If rsUpload.State = ADODB.ObjectStateEnum.adStateOpen Then rsUpload.Close()
            strString = " Select ScheduleNo,ScheduleDate,SchStartDate,SchEndDate,SchCustUpdDt From Sched_Smiel_Mst Where Unit_Code='" + gstrUNITID + "' And Customer_Code = '" & stracccode & "' and ScheduleNo = '" & strschno & "'"
            rsUpload.Open(strString, mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If Not rsUpload.EOF Then
                strString = "Schedule Already Uploaded With Following Information " & vbCrLf & "" & vbCrLf & "DS No.     " & rsUpload.Fields("ScheduleNo").Value & vbCrLf & "Schedule Date " & rsUpload.Fields("ScheduleDate").Value & vbCrLf & "Schedule Start Date " & rsUpload.Fields("SchStartDate").Value & vbCrLf & "Schedule End Date   " & rsUpload.Fields("SchEndDate").Value & vbCrLf & "Schedule Customer Updated Date " & rsUpload.Fields("SchCustUpdDt").Value & vbCrLf & "" & vbCrLf & "You are Uploading New DS with Following Information " & vbCrLf & vbCrLf & "DS No.     " & strschno & "" & vbCrLf & "Schedule Date " & strMstFields(3) & "" & vbCrLf & "Schedule Start Date " & strMstFields(5) & vbCrLf & "Schedule End Date   " & strMstFields(6) & vbCrLf & "Schedule Customer Updated Date " & strCurDtSQL & vbCrLf & vbCrLf & "               Do You Want To Upload it Again "
                If MsgBox(strString, MsgBoxStyle.YesNo, "empower") = MsgBoxResult.Yes Then
                    blnUpdfg = True
                Else
                    DailyScheduleEntry = "N»"
                    Exit Function
                End If
            End If ' End if For rsUpload
        End If
        strsql = "Set Dateformat 'dmy' " & vbCrLf
        strDeleteScheduleHistory = "Set Dateformat 'dmy' " & vbCrLf
        Dim intDayOfWeek As Short
        intDayOfWeek = Weekday(GetServerDate, FirstDayOfWeek.Monday)
        If blnUpdfg = True Then
            strString = " DS Upload will Overwrite Existing Data.Do u Want to Countinue "
            If MsgBox(strString, MsgBoxStyle.YesNo, "empower") = MsgBoxResult.No Then
                DailyScheduleEntry = "N»"
                Exit Function
            End If
            strSched_Smiel_mst = strSched_Smiel_mst & " Update Dailymktschedule Set Status = 0,RevisionNo = RevisionNo + 1 Where Unit_Code='" + gstrUNITID + "' And Account_Code = '" & stracccode & "' and DSNO = '" & strschno & "'" & vbCrLf
            strSched_Smiel_mst = strSched_Smiel_mst & " Update Dailymktschedule_Tentative Set Status = 0,RevisionNo = RevisionNo + 1 Where Unit_Code='" + gstrUNITID + "' And Account_Code = '" & stracccode & "' and DSNO = '" & strschno & "'" & vbCrLf
            strSched_Smiel_mst = strSched_Smiel_mst & " Insert Into Sched_Smiel_Mst (Customer_Code, ScheduleNo, ScheduleDate, SchStartDate, SchEndDate, SchUpdDate, "
            strSched_Smiel_mst = strSched_Smiel_mst & " SchCustUpdDt, SchEntryDate,Control_No,WeekDayOfSchedule,Ent_Dt,Unit_Code) Values ( '" & stracccode & "', '"
            strSched_Smiel_mst = strSched_Smiel_mst & strMstFields(2) & "', '" & strMstFields(3) & "', '"
            strSched_Smiel_mst = strSched_Smiel_mst & strMstFields(5) & "', '" & strMstFields(6) & "', "
            strSched_Smiel_mst = strSched_Smiel_mst & " '" & strCurDtSQL & "' ,  '" & strFileDtTm & "','" & strCurDtSQL & "'," & strMstFields(8) & "," & intDayOfWeek & ",getdate(),'" + gstrUNITID + "')" & vbCrLf
        Else
            'Update Master Table
            strSched_Smiel_mst = ""
            strSched_Smiel_mst = strSched_Smiel_mst & " Insert Into Sched_Smiel_Mst (Customer_Code, ScheduleNo, ScheduleDate, SchStartDate, SchEndDate, SchUpdDate, "
            strSched_Smiel_mst = strSched_Smiel_mst & " SchCustUpdDt, SchEntryDate,Control_No,WeekDayOfSchedule ,Ent_Dt,Unit_Code) Values ( '" & stracccode & "', '"
            strSched_Smiel_mst = strSched_Smiel_mst & strMstFields(2) & "', '" & strMstFields(3) & "', '"
            strSched_Smiel_mst = strSched_Smiel_mst & strMstFields(5) & "', '" & strMstFields(6) & "', "
            strSched_Smiel_mst = strSched_Smiel_mst & " '" & strCurDtSQL & "' ,  '" & strFileDtTm & "','" & strCurDtSQL & "'," & strMstFields(8) & "," & intDayOfWeek & ",getdate() ,'" + gstrUNITID + "')" & vbCrLf
        End If
        'Detail Part
        mP_Connection.BeginTrans()
        mP_Connection.Execute("delete from dailymktSchedule_Temp where unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        'Execute SQL Statment of Schedule_Smiel_mst
        If Len(Trim(strSched_Smiel_mst)) > 0 Then
            mP_Connection.Execute(strSched_Smiel_mst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        If blnUpdfg = True Then
            mP_Connection.Execute("Delete from dailyschedule_history where schedule_no = '" & strschno & "' and  customer_code = '" & stracccode & "' and  unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        mP_Connection.CommitTrans()
        While Len(strBuffer) > 0
            System.Windows.Forms.Application.DoEvents()
            counter = 0 : intRowNo = intRowNo + 1
            intRowSeperator = InStr(strBuffer, "^")
            strRow = Mid(strBuffer, 1, intRowSeperator - 1) 'Get 1 Row at a Time
            'Seperate Fields 1.Customer PO Number  2.Customer Part Number 3. Transaction start date 4. Transaction end date 5. Schedule quantity 6.Dispatch Quantity
            While Len(strRow) > 0
                counter = counter + 1 : intPos = InStr(strRow, "|")
                If intPos > 0 Then
                    strFields(counter) = Mid(strRow, 1, intPos - 1) 'Get Individual Fields in ARRAY
                Else
                    strFields(counter) = strRow
                    strRow = ""
                End If
                strRow = Mid(strRow, intPos + 1)
            End While
            'Check Mismatched/ new part no.
            strSbuItCode = strFields(2)
            If Not ValidateData("CD.Cust_DrgNo", " Cust_Ord_Dtl CD,Cust_Ord_Hdr CH ", " CD.unit_Code = CH.unit_Code and CD.Account_Code = CH.Account_Code AND CD.Cust_Ref = CH.Cust_Ref AND CD.Amendment_No = CH.Amendment_No AND  CD.Account_Code = '" & stracccode & "' AND CD.Cust_DrgNo = '" & strSbuItCode & "' AND CD.ACTIVE_Flag = 'A' AND PO_Type = 'O' and CD.UNIT_code='" & gstrUNITID & "'") Then
                'Customer Drwg No Not Found
                MsgBox("Customer Drawing Number " & strSbuItCode & " not found." & vbCr & "Adding this Drawing No. in JunkFile." & vbCr & "Continuing with rest of the Schedule File...", MsgBoxStyle.Information, "empower")
                a.WriteLine((" Customer Drawing No. - " & strSbuItCode))
                xmlBasic.WriteLine(("<Cust_DrgNo>" & strSbuItCode & "</Cust_DrgNo>"))
                blnJunkgenerate = True
                lnRowCounter = lnRowCounter - 1
                'Line (Output,nInSchFile, strBuffer)
                GoTo ItemCodeNotFound
            Else
                'Get Internal Item Code
                strItemCode = Trim(SelectDataFromTableNew("Item_Code", " Cust_Ord_Dtl CD,Cust_Ord_Hdr CH ", " CD.unit_Code = CH.unit_Code and  CD.Account_Code = CH.Account_Code AND CD.Cust_Ref = CH.Cust_Ref AND CD.Amendment_No = CH.Amendment_No AND  CD.Account_Code = '" & stracccode & "' AND CD.Cust_DrgNo = '" & strSbuItCode & "' AND CD.ACTIVE_Flag = 'A' AND PO_Type = 'O' and cd.UNIT_code='" & gstrUNITID & "'"))
            End If
            System.Windows.Forms.Application.DoEvents()
            intYYYYMM = CInt(VB6.Format(strFields(3), "yyyymm")) 'YYYYMM Format
            'If Item has monthly Schedule then it will be updated with status 0
            If ValidateData("Item_Code", "MonthlyMKTSchedule", " Account_Code = '" & stracccode & "' AND Cust_Drgno = '" & strSbuItCode & "' AND Item_Code = '" & strItemCode & "' AND Status = 1 AND UNIT_code='" & gstrUNITID & "' and Year_Month = " & intYYYYMM) Then
                If rsUpload.State = ADODB.ObjectStateEnum.adStateOpen Then rsUpload.Close()
                rsUpload.Open(" Select Despatch_qty From MonthlymktSchedule Where Unit_Code='" + gstrUNITID + "' And Account_Code = '" & stracccode & "' AND Cust_Drgno = '" & strSbuItCode & "' AND Item_Code = '" & strItemCode & "' AND Status = 1 AND Year_Month = " & intYYYYMM & " and Despatch_qty > 0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If Not rsUpload.EOF Then
                    intErrorNumber = intErrorNumber + 1
                    strErrorMessage = strErrorMessage & " For Item Code " & strItemCode & "  Despatch Already Exist in Monthly Schedule With Quantity " & rsUpload.Fields("despatch_qty").Value & ".Please Varify Data and Upload it Again" & vbCrLf
                    DailyScheduleEntry = "N»" & strErrorMessage
                    Exit Function
                End If
                strMonthlymktSchedule = strMonthlymktSchedule & vbCrLf & vbCrLf & " Update MonthlyMKTSchedule Set Status = 0 Where Unit_Code='" + gstrUNITID + "' And Account_Code = '" & stracccode & "'" & " And Cust_DrgNo = '" & strSbuItCode & "' AND Item_Code = '" & strItemCode & "'" & " AND Status = 1 AND Year_Month = " & intYYYYMM
                'Move to Next Item In List
            End If
            System.Windows.Forms.Application.DoEvents()
            If strFields(7) = "F" Then
                'Insert Script for DailymktSchedule
                strDailymktSchedule = "  Insert into DailymktSchedule_temp (Account_Code,Trans_date,Item_code, Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty, Status,RevisionNo,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DSDateTime,DSNO,Unit_Code)" & "Values('" & stracccode & "','" & strFields(3) & "','" & strItemCode & "'," & "'" & strSbuItCode & "',1,isnull('" & Val(strFields(5)) & "',0),0," & "1,0,getDate() ,'PORT',getdate(),'PORT','" & strFileDtTm & "','" & strschno & "','" + gstrUNITID + "')"
                mP_Connection.BeginTrans()
                mP_Connection.Execute(strDailymktSchedule, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.CommitTrans()
                lnInsertCounter = lnInsertCounter + 1
            Else
                'Insert Script For DailymktSchedule_Tentative
                strDailymktSchedule = " Insert into DailymktSchedule_Tentative (Account_Code,Trans_date,Item_code," & "Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty, " & "Status,RevisionNo,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DSDateTime,DSNO,Unit_Code)" & "Values('" & stracccode & "','" & strFields(3) & "','" & strItemCode & "'," & "'" & strSbuItCode & "',1,isnull('" & Val(strFields(5)) & "',0),0," & "1,0,getDate() ,'PORT','" & strCurDtSQL & "','PORT','" & strFileDtTm & "','" & strschno & "','" + gstrUNITID + "')"
                mP_Connection.BeginTrans()
                mP_Connection.Execute(strDailymktSchedule, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.CommitTrans()
            End If
            'If Schedule already uploaded. then an entry will be deleted from dailyschedule_history
            strDailySchedule_History = "  Insert into DailySchedule_History (ScheduleCode ,Customer_code,Schedule_no,SchMade_dt,Schedule_date,Schedule_type," & " Cust_drgno,Item_code,Schedule_qty,Ent_dt,Ent_userid,Upd_dt,Upd_userid,Unit_Code)" & " Values('" & strMstFields(4) & "','" & stracccode & "','" & strschno & "','" & strMstFields(3) & "','" & strFields(3) & "','" & strFields(7) & "'" & ",'" & strSbuItCode & "','" & strItemCode & "'," & strFields(5) & ",getdate(),'PORT',getdate(),'PORT','" + gstrUNITID + "')"
            mP_Connection.BeginTrans()
            mP_Connection.Execute(strDailySchedule_History, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
ItemCodeNotFound:  'In Case Item is Not Found it is SKIPPED
            If InStr(strBuffer, "^") > 0 Then
                strBuffer = Mid(strBuffer, InStr(strBuffer, "^") + 1)
            End If
            If Len(strBuffer) = 0 Or strBuffer = "^" Then strBuffer = ""
        End While
        If blnUpdfg = True Then
            '        strSQL = strSQL & vbCrLf & " Update DailymktSchedule Set Despatch_qty = isnull(( Select Despatch_qty from dailymktschedule a where DSNO = '" & strschno & "' and account_Code = '" & stracccode & "' and Revisionno = 1 and item_Code = dailymktschedule.Item_code ),0) where DSNO = '" & strschno & "' and account_Code = '" & stracccode & "' and status = 1 and revisionno = 0"
            strDailymktUpdate = " Update a set a.Despatch_Qty = b.Despatch_Qty from DailyMktSchedule_temp a,"
            strDailymktUpdate = strDailymktUpdate & " DailyMktSchedule b Where A.Unit_Code=B.Unit_code And  A.Unit_Code='" + gstrUNITID + "' And  a.Cust_DrgNo = b.Cust_DrgNo And a.Account_Code = b.Account_Code and a.Trans_date = b.Trans_date"
            strDailymktUpdate = strDailymktUpdate & " And a.DSNo = b.DSNo and a.revisionNo =0 and b.revisionNo =1 and a.DSNO = '" & strschno & "'"
            strDailymktUpdate = strDailymktUpdate & " and a.account_Code = '" & stracccode & "' and a.Status =1"
        Else
            strDailymktUpdate = ""
        End If
        mP_Connection.BeginTrans()
        'Execute SQL Statment for MonthlymktSchedule
        If Len(Trim(strMonthlymktSchedule)) > 0 Then
            mP_Connection.Execute(strMonthlymktSchedule, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        'Execute SQL Statment for DailymktSchedule
        'Execute SQL Statment for Update DailySchedule
        If Len(Trim(strDailymktUpdate)) > 0 Then
            mP_Connection.Execute(strDailymktUpdate, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        mP_Connection.Execute("Insert into DailymktSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,RevisionNo,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DSDateTime,DSNO,Unit_code) Select Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,RevisionNo,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DSDateTime,DSNO,Unit_Code from DailymktSchedule_temp  where unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If rsUpload.State = ADODB.ObjectStateEnum.adStateOpen Then rsUpload.Close()
        rsUpload = mP_Connection.Execute("Select Count(*) from DailyMktSchedule_temp Where Unit_Code='" + gstrUNITID + "'")
        If blnUpdfg = False Then
            If Not rsUpload.EOF Then
                If lnInsertCounter <> rsUpload.Fields(0).Value Then
                    DailyScheduleEntry = "N»No. of Rows of Firm Schedule in Text File is Not Equal to No. of Rows Inserted In DataBase"
                    mP_Connection.RollbackTrans()
                    mP_Connection.Execute(" Delete from Dailyschedule_history where schedule_no = '" & strschno & "' and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute(" Delete from DailymktSchedule_Tentative where DSNO = '" & strschno & "' and status = 1  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    Exit Function
                End If
            Else
                DailyScheduleEntry = "N» Number of Rows Inserted In Database is equal to Zero"
                mP_Connection.RollbackTrans()
                mP_Connection.Execute(" Delete from Dailyschedule_history where schedule_no = '" & strschno & "' and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(" Delete from DailymktSchedule_Tentative where DSNO = '" & strschno & "' and status = 1  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Exit Function
            End If
        End If
        If rsUpload.State = ADODB.ObjectStateEnum.adStateOpen Then rsUpload.Close()
        rsUpload = mP_Connection.Execute("Select item_Code,Trans_date from DailymktSchedule Where Unit_Code='" + gstrUNITID + "' And Status = 1 and dsno = '" & strschno & "' and despatch_qty > schedule_quantity")
        If Not rsUpload.EOF Then
            Intcounter = 1
            strString = "N»DS Could not be uploaded due to Despatch Quantity is Excees then New Schedule Quantity for Following Item(s)" & vbCrLf & vbCrLf
            strString = strString & "   Item Code                            " & " Transaction Date " & vbCrLf
            strsql = ""
            strDeleteScheduleHistory = ""
            While Not rsUpload.EOF
                strsql = strsql & Intcounter & ". " & rsUpload.Fields("Item_Code").Value & "                       " & rsUpload.Fields("Trans_date").Value & vbCrLf
                Intcounter = Intcounter + 1
                rsUpload.MoveNext()
            End While
            mP_Connection.RollbackTrans()
            strString = strString & strsql
            DailyScheduleEntry = strString
            rsUpload.Close()
            Exit Function
        End If
        mP_Connection.CommitTrans()
        System.Windows.Forms.Application.DoEvents()
        '---------------------------------------
        xmlBasic.WriteLine(("</Junk>"))
        xmlBasic.Close() 'Close File
        a.Close()
        '---------------------------------------
        System.Windows.Forms.Application.DoEvents()
        'Check whether a file with the same name already exists
        If Len(Dir(strSchPath & "Loaded Files\" & strSchFILE)) > 0 Then
            'Remove existing file
            Kill(strSchPath & "\Loaded Files\" & strSchFILE)
        End If
        Dim ass As Scripting.FileSystemObject
        ass = New Scripting.FileSystemObject
        ass.MoveFile(strSchPath & strSchFILE, strSchPath & "Loaded Files\" & strSchFILE)
        If blnJunkgenerate = True Then
            ass.MoveFile(strSchPath & "JunkFile.jnk", strSchPath & "Loaded Files\" & VB.Left(strSchFILE, Len(strSchFILE) - 4) & "JunkFile.jnk")
        End If
        DailyScheduleEntry = "Y»File has been uploaded successfully !"
        ass = Nothing
        ass2 = Nothing
        ass3 = Nothing
        Exit Function
Errorhandler:
        mP_Connection.RollbackTrans()
        mP_Connection.Execute(" Delete from Dailyschedule_history where schedule_no = '" & strschno & "' and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(" Delete from DailymktSchedule_Tentative where DSNO = '" & strschno & "' and status = 1  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ValidateControlNo(ByVal strScheduleNo As String, ByVal StrCustomerCode As String, ByVal lngControlNo As Integer) As String
        On Error GoTo Errorhandler
        Dim rsCon As New ADODB.Recordset
        Dim lngCounter As Integer
        Dim lngCounter1 As Integer
        Dim strString As String
        Dim blnCheckExistence As Boolean
        Dim arr() As String
        If rsCon.State = ADODB.ObjectStateEnum.adStateOpen Then rsCon.Close()
        rsCon.Open(" Select Distinct ScheduleNo from Sched_Smiel_mst Where Unit_Code='" + gstrUNITID + "' And Customer_Code = '" & StrCustomerCode & "' and Control_No = " & lngControlNo & " and ScheduleNo not in ('" & strScheduleNo & "') and Control_no > 0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not rsCon.EOF Then
            ValidateControlNo = "N»Control Number Already Exist For Schedule No:- " & rsCon.Fields("ScheduleNo").Value
            rsCon = Nothing
            Exit Function
        End If
        ValidateControlNo = "Y»"
        rsCon = Nothing
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub ctlUploadSchedulesHDR_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlUploadSchedulesHDR.Click
        On Error GoTo ErrHandler
        MsgBox("No HELP available for this Screen.", MsgBoxStyle.Information, "empower")
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdUploadSchd_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles cmdUploadSchd.ButtonClick
        On Error GoTo ErrHandler

        If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then Me.Close() : Exit Sub
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT  'Called From Upload Schedules
                Call TransferCheckedSchedules()
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstStatus_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lstStatus.ItemChecked
        On Error GoTo ErrHandler
        Dim Item As System.Windows.Forms.ListViewItem = lstStatus.Items(e.Item.Index)
        Dim intCount As Integer
        If Me.lstStatus.Items.Item(Item.Index).Checked = False Then
            While intCount <> Me.lstStatus.Items.Count
                lstStatus.Items.Item(intCount).Checked = False
                intCount = intCount + 1
            End While
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
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
End Class