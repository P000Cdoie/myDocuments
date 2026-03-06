Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0029
	Inherits System.Windows.Forms.Form
	'*****************************************************************************
	' Copyright©2002             - MIND
	' Form Name                  - frmMKTTRN0029.frm
	' Created by                 - Jasmeet Singh Bawa
	' Created Date               - 20/02/2004
	' Form Description           - Upload Schedules into Database
	' Changes Done by Nisha on 19-oct-2004 for capturing tentative data into forecast
	' Revision Date              -   08-03-2005
	' Revision History           -   Changed for uploading of schedule for the Active Items and Authorised
	'SO in case there are more than one Item for the combination of
	'Account Code and Cust_DrgNo
	' Form Description           -   Upload Schedules into Database
	'*****************************************************************************
	' Revision Arshad            -   Arshad Ali
	' Revision Date              -   05-04-2005
	' Revision History           -   There was a bug when uploading E-nagare schedule
	'                                "Flag" values were inserted in item code column in database
	'                                if the value of eNagareUploadingOnBasisOfSO is 0 in Sales_parameter table
	'-----------------------------------------------------------------------------
	' Revision Arshad            -   Sandeep Chadha
	' Revision Date              -   26-Apr-2005
	' Revision History           -   In case of DISchedule Data should be also save in ForeCast_Mst
	'-----------------------------------------------------------------------------
	' Revision History           -   Sandeep Chadha
	' Revision Date              -   06-June-2005
	' Revision History           -   In Case of Multiple Records in CustItem_Mst
	'                            -   Item Code is Saved as False resulting Uploading Schedule is not shown in Front-End.
	'-----------------------------------------------------------------------------
	' Revision By                -   Prashant Dhingra
	' Revision Date              -   21-Jun-2005
	' Revision History           -   Schedule Type added to eliminate duplicate entries
	'----------------------------------------------------------------------------------------------
	' Revision Date              -   27/03/2006
	' Revision By                -   Davinder Singh
	' Issue ID                   -   17378
	' Revision History           -   Changes to send the data in the Forecast in case of DI Spares
	'----------------------------------------------------------------------------------------------
	' Revision Date              -   24/04/2006
	' Revision By                -   Davinder Singh
	' Issue ID                   -   17628
	' Revision History           -   Primary Key Violation error occurs sometimes during Uploading DI Schedules
	'------------------------------------------------------------------------------------------------------------------------------------------
	' Revision Date             -   02/06/2006
	' Revision By               -   Davinder Singh
	' Issue ID                  -   17995
	' Revision History          -   To also check the newly added Po_Type='M' during SO checking in DISpares function
	'------------------------------------------------------------------------------------------------------------------------------------------
	' Revision Date             -   31/08/2006
	' Revision By               -   Davinder Singh
	' Issue ID                  -   18532
	' Revision History          -   During uploading DI Schedule an error message 'BOF or EOF true' comes due to not checking of
	'                               EOF on the opened recordset.
	'------------------------------------------------------------------------------------------------------------------------------------------
	' Revision Date             -   08/09/2006
	' Revision By               -   Ashutosh Verma
	' Issue ID                  -   18573
	' Revision History          -    Upload maruti's new format for query upload.
	'------------------------------------------------------------------------------------------------------------------------------------------
	' Revision Date             -   02 Jan 2008
	' Revision By               -   Manoj Kumar Vaish
	' Issue ID                  -   22035
	' Revision History          -   Wrong query in Upload maruti's new format for query upload function.
	'------------------------------------------------------------------------------------------------------------------------------------------
	'Revision Date      :         08 Feb 2008
	'Revised By         :         Prashant Rajpal
	'Revision for       :        Issue No.22228- Schedule uploading (only tentative schedule is uploaded )
	'------------------------------------------------------------------------------------------------------------------------------------------
	'Revision Date      :         02 June 2008
	'Revised By         :         Prashant Rajpal
	'Revision for       :        As per the RFC - Now the DI Schedule is alos uploaded in Daily kt schedule , if E-nagare is not THERE
    '*************************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090223-27780
    'Revision Date      : 25 Feb 2009
    'History            : Problem while uploading DI Spares file
    '***********************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090416-30261
    'Revision Date      : 16 Apr 2009
    'History            : To check the 'O' type PO while uploading the OESmaruti Schedule from DIQuery Option
    '********************************************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090424-30570
    'Revision Date      : 24 Apr 2009
    'History            : SO type was not updating while uploading the maruti SPD Schedule through DI Query Option
    '********************************************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090505-31005
    'Revision Date      : 05 May 2009
    'History            : While Uploading the  Maruti Schedule through DI Query option user will select SO type.
    '********************************************************************************************************
    'Modified by Amit Rana on 06/May/2011 for multiunit change
    '********************************************************************************************************
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10238894 
    'Revision Date      : 20 june 2012
    'History            : Time for spare and diquery changed by 11:59 instead of 23:59
    '********************************************************************************************************
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10263152  
    'Revision Date      : 14 Aug 2012
    'History            : changes in Schedule Uploading -mul 
    '********************************************************************************************************
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10343323   
    'Revision Date      : 12 Feb 2013
    'History            : changes in : customer item master active check integrated .
    '********************************************************************************************************
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10565410  
    'Revision Date      : 26 mar 2014
    'History            : Error resolved during spare file uploading :giving string conversion error
    '********************************************************************************************************
    'Created By     : Parveen Kumar
    'Created On     : 13 FEB 2015
    'Description    : eMPro Vehicle BOM
    'Issue ID       : 10737738 
    '----------------------------------------------------------------------------------

    Private mlngFormTag As Short 'Form Tag
    Dim m_intFilterIndex As Short
    Dim mstrHelp() As String
    Dim SchUpdFlag As Boolean = False    '10737738
    Dim schkanbanunitFlag As Boolean = False    '10737738


    Private Sub cmdclose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        On Error GoTo ErrHandler
        Me.Close() 'Unload the Form
        Exit Sub
        'Execution of Error Handler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'Changes against 10737738 
    Private Sub ChkVBSchUpdFlag()
        Dim strSql As String = String.Empty

        Try

            strSql = " select top 1 1 from sales_parameter where Unit_Code='" & gstrUNITID & "' and SCHEDULE_UPLOAD_CONFIG = 1  "
            SchUpdFlag = IsRecordExists(strSql)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub ChkKanbannunit()
        Dim strSql As String = String.Empty

        Try

            strSql = " select top 1 1 from FGB_CONFIG_MST where Unit_Code='" & gstrUNITID & "'"
            schkanbanunitFlag = IsRecordExists(strSql)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub cmdcustcode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCode.Click

        Try
            'Changes against 10737738 
            If SchUpdFlag = True Then
                Call ShowCode_Desc("SELECT customer_code,cust_name FROM customer_mst  Where Unit_Code='" + gstrUNITID + "' and SCH_UPLOAD_CODE='NAGARE' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))", TxtCustCode, LblCustDesc)
            Else
                Call ShowCode_Desc("SELECT customer_code,cust_name FROM customer_mst  Where Unit_Code='" + gstrUNITID + "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))", TxtCustCode, LblCustDesc)
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub cmdFileSelector_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFileSelector.Click
        cdlgFileSelectorOpen.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        cdlgFileSelectorOpen.ShowDialog()
        txtDBFFilePath.Text = cdlgFileSelectorOpen.FileName
        txtDBFFilePath.ForeColor = System.Drawing.Color.Black
        m_intFilterIndex = cdlgFileSelectorOpen.FilterIndex
    End Sub

    Private Sub cmdTransfer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTransfer.Click
        On Error GoTo ErrHandler
        Dim strPOtype As String
        '----------Calling the function to upload E Nagare Schedules
        If Len(Trim(txtDBFFilePath.Text)) = 0 Then
            MsgBox("File is Invalid or has not been entered", MsgBoxStyle.Information, ResolveResString(100))
            txtDBFFilePath.Focus()
            Exit Sub
        End If
        If Len(Trim(TxtCustCode.Text)) = 0 Then
            MsgBox("Please enter the Customer Code", MsgBoxStyle.Information, ResolveResString(100))
            TxtCustCode.Focus()
            Exit Sub
        End If
        If OptDI.Checked = True Then
            Call DISchedules()
        ElseIf OptNagare.Checked = True Then
            Call UpdateDAILYSchedule((txtDBFFilePath.Text))
        ElseIf optDIQuery.Checked = True Then
            If Trim(cmbsotype.Text) = "" Then
                MessageBox.Show("Select SO Type before uploading the schedule.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                cmbsotype.Focus()
                Exit Sub
            Else
                If UCase(Trim(cmbsotype.Text)) = "MRP-SPARES" Then
                    strPOtype = "M"
                Else
                    strPOtype = "O"
                End If
                Call UpdateDIQuerySchedule(strPOtype)
            End If
        Else
            Call DISpares()
        End If
        Exit Sub
        'Execution of Error Handler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function UpdateDIQuerySchedule(ByVal pstrPOtype As String) As Object
        '-------------------------------------------------------------------------------------------------------------------------------------------
        ' Revision Date             -   08/09/2006
        ' Revision By               -   Ashutosh Verma
        ' Issue ID                  -   18573
        ' Revision History
        '------------------------------------------------------------------------------------------------------------------------------------------
        ' Revision Date             -   08 Jan 2008
        ' Revision By               -   Manoj Kr. Vaish
        ' Issue ID                  -   22035
        ' Revision History          -   To check the Spare Po_Type='M' in place of OEM during SO checking
        '------------------------------------------------------------------------------------------------------------------------------------------
        'Revised By                 -   Manoj Kr. Vaish
        'Issue ID                   -   eMpro-20090416-30261
        'Revision Date              -   16 Apr 2009
        'History                    -   To check the 'O' type PO while uploading the OES Maruti Schedule from DIQuery Option
        '********************************************************************************************************
        'Revised By                 -   Manoj Kr. Vaish
        'Issue ID                   -   eMpro-20090505-31005
        'Revision Date              -   05 May 2009
        'History                    -   To check the PO Type according to the parameter passed in this function
        '********************************************************************************************************
        On Error GoTo ErrHandler
        Dim FSODISpares As Scripting.FileSystemObject
        Dim FSODISparesReadStatus As Scripting.TextStream
        Dim strstatus As String
        Dim i As Short
        Dim dblqty As Double
        Dim strMasterString As String
        Dim ArrMasterArray() As String
        Dim ArrSplitData() As String
        Dim stracccode As String
        Dim strItemCode As String
        Dim strSbuItCode As String
        Dim intYYYYMM As Integer
        Dim strsql As String
        Dim strSQLA As String
        Dim dblDispatchqty As Double
        Dim dblPrevSchedQty As Double
        Dim RsObjInsert As ADODB.Recordset
        Dim RsObjQuery As ADODB.Recordset
        Dim Rs As ADODB.Recordset
        Dim intMaxRNo As Short
        Dim strPOType As String
        Dim strcustdrgno As String
        Dim strquantity As String
        Dim strUNLOC As String
        Dim StrUSLOC As String
        Dim strKanbanNo As String
        Dim strschdate As String
        Dim strpricechange As String
        Dim strbatchcode As String
        Dim strprice As String
        Dim strvendorcode As String

        FSODISpares = New Scripting.FileSystemObject
        RsObjInsert = New ADODB.Recordset
        RsObjQuery = New ADODB.Recordset
        Rs = New ADODB.Recordset
        On Error GoTo ErrHandler
        mP_Connection.BeginTrans()
        FSODISparesReadStatus = FSODISpares.OpenTextFile(txtDBFFilePath.Text, Scripting.IOMode.ForReading, False)
        ''----Delete all data from temporary table Tmp_Enagarodtl for user's IP
        mP_Connection.Execute("DELETE FROM Tmp_Enagarodtl Where Unit_Code='" + gstrUNITID + "' AND Session_id='" & gstrIpaddressWinSck & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If pstrPOtype = "O" Then


            ''----Read the data from text file and put it into the temporary table Tmp_Enagarodtl
            While Not FSODISparesReadStatus.AtEndOfLine
                strMasterString = ""
                strstatus = FSODISparesReadStatus.ReadLine()
                ArrSplitData = Split(strstatus, " ")
                For i = 0 To UBound(ArrSplitData)
                    If Len(Trim(ArrSplitData(i))) > 0 Then
                        strMasterString = strMasterString & ArrSplitData(i) & "»"
                    End If
                Next
                ArrSplitData = Split(strMasterString, "»")
                If UBound(ArrSplitData) >= 9 Then
                    strMasterString = ""
                    For i = 1 To UBound(ArrSplitData)
                        If i < 5 Or i > UBound(ArrSplitData) - 4 Then
                            If Len(Trim(ArrSplitData(i))) > 0 Then
                                strMasterString = strMasterString & ArrSplitData(i) & "»"
                            End If
                        End If
                    Next
                End If
                'NAGARE and COMP are considered as two seperate words .So they are concatenated and inserted into the table
                ArrMasterArray = Split(strMasterString, "»")
                If UBound(ArrMasterArray) > 6 Then
                    If IsDate(ArrMasterArray(0)) Then
                        If Len(ArrMasterArray(2)) = 5 Then
                            mP_Connection.RollbackTrans()
                            MsgBox("Invalid Schedule option Selected. File is E Nagare Schedule.", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Function
                        Else
                            mP_Connection.Execute("INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Freq,UNIT_CODE) values('" & gstrIpaddressWinSck & "','" & Trim(ArrMasterArray(1)) & "','" & Trim(ArrMasterArray(2)) & "' ,'" & Trim(ArrMasterArray(3)) & "','" & Trim(ArrMasterArray(4)) & "',' ','" & Trim(ArrMasterArray(5)) & "','" & Trim(ArrMasterArray(0)) & "','11:59','1','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    End If
                End If
            End While
        Else
            While Not FSODISparesReadStatus.AtEndOfLine
                strMasterString = ""
                strstatus = FSODISparesReadStatus.ReadLine()
                strvendorcode = Mid(strstatus, 46, 7)
                strcustdrgno = Mid(strstatus, 54, 15)
                strquantity = Mid(strstatus, 69, 9)
                strUNLOC = Mid(strstatus, 78, 9)
                StrUSLOC = Mid(strstatus, 96, 7)
                strKanbanNo = Mid(strstatus, 104, 14)
                strschdate = Mid(strstatus, 11, 17)
                strpricechange = Mid(strstatus, 138, 21)
                strbatchcode = Mid(strstatus, 160, 5)
                strprice = Mid(strstatus, 165, 12)

                If IsDate(Mid(strstatus, 11, 11)) = True Then
                    mP_Connection.Execute("INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Freq,UNIT_CODE,price_change_flag,batch_code,price) values('" & gstrIpaddressWinSck & "','" & strvendorcode & "','" & strcustdrgno & "' ,'" & strquantity & "','" & strUNLOC & "','" & StrUSLOC & "','" & strKanbanNo & "','" & strschdate & "','11:59','1-1','" + gstrUNITID + "','" & strpricechange & "' ,'" & strbatchcode & "','" & strprice & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If

            End While
        End If
        FSODISpares = Nothing
        FSODISparesReadStatus = Nothing
        ''----Fetch the whole data from temporary table for current IP to in a Recordset
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RsObjInsert.Open("SELECT * FROM Tmp_enagarodtl Where Unit_Code='" + gstrUNITID + "' And Session_ID='" & gstrIpaddressWinSck & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                If Rs.State = ADODB.ObjectStateEnum.adStateOpen Then Rs.Close()
                'To retrieve Customer code line by line
                stracccode = ""
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                RsObjQuery.Open("SELECT customer_code FROM customer_mst Where Unit_Code='" + gstrUNITID + "' And cust_vendor_code='" & Trim(RsObjInsert.Fields("vendor_code").Value) & "' and customer_code = '" & Trim(Me.TxtCustCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjQuery.EOF Then
                    stracccode = Trim(RsObjQuery.Fields("customer_code").Value)
                Else
                    MsgBox("No Data found in the Customer Master for the combination of seleted Customer Code[" & Trim(TxtCustCode.Text) & "] and customer vendor code[" & Trim(RsObjInsert.Fields("vendor_code").Value) & "] available in the file.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                ''----Pick item code from custitem_mst
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                If pstrPOtype = "O" Then
                    RsObjQuery.Open("SELECT item_code FROM custitem_mst Where Unit_Code='" + gstrUNITID + "' And cust_drgno='" & Trim(RsObjInsert.Fields("cust_drgno").Value) & "' AND account_code='" & stracccode & "' and active=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                Else
                    RsObjQuery.Open("SELECT item_code FROM custitem_mst Where Unit_Code='" + gstrUNITID + "' And cust_drgno='" & strcustdrgno & "' AND account_code='" & stracccode & "' and active=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                End If

                If Not RsObjQuery.EOF Then
                    strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                End If
                If RsObjQuery.RecordCount > 1 Then
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("SELECT D.Item_code,H.PO_TYPE from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE=D.UNIT_CODE And H.Unit_Code='" + gstrUNITID + "' And H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type ='" & Trim(pstrPOtype) & "' and D.Active_Flag='A' and D.cust_drgNo='" & strcustdrgno & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If Not RsObjQuery.EOF Then
                        strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                        strPOType = Trim(RsObjQuery.Fields("Po_Type").Value)
                        GoTo Onerec
                    Else
                        If MsgBox(" There are more than 1 item code defined for this Customer part Code : " & Trim(RsObjInsert.Fields("cust_drgno").Value) & "." & vbCrLf & " Proceed with it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                            GoTo Onerec
                        Else
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        End If
                    End If
                ElseIf RsObjQuery.RecordCount < 1 Then  'Message for Item code is not Active and roll back the uploading
                    MsgBox(" Item Code not found for Customer Part Code code : " & Trim(RsObjInsert.Fields("cust_drgNo").Value) & vbCrLf & " Please correct the data first. It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                Else
Onerec:
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("select eNagareUploadingOnBasisOfSO FROM sales_parameter Where Unit_Code='" + gstrUNITID + "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If RsObjQuery.Fields(0).Value = True Then 'Value is set for eNagareUploadingOnBasisOfSO in sales_parameter
                        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                        RsObjQuery.Open("SELECT D.Item_code,H.PO_TYPE from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE=D.UNIT_CODE AND H.Unit_Code='" + gstrUNITID + "' AND H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 AND H.po_type='" & Trim(pstrPOtype) & "' AND D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If RsObjQuery.RecordCount = 0 Then 'There is no SO Active & Authorized
                            MsgBox(" There is no SO Authorized or Active for Cust Part Code: " & Trim(RsObjInsert.Fields("cust_drgNo").Value) & " for selected Customer. " & vbCrLf & " It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        Else
                            strItemCode = Trim(RsObjQuery.Fields("Item_Code").Value)
                            strPOType = Trim(RsObjQuery.Fields("Po_Type").Value)
                        End If
                    End If
                End If
                '----If current kanban already exist then read its Qty. and Reduce it from the respective quantities of DailyMKTSchedule and delete from mkt_enagaredtl
                strSQLA = "select Quantity from mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' AND Account_code = '" & stracccode & "' and Item_code = '" & strItemCode & "' and Cust_drgno = '" & RsObjInsert.Fields("cust_drgno").Value & "' and kanbanno = '" & RsObjInsert.Fields("kanbanno").Value & "'"
                Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                Rs.Open(strSQLA, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Rs.RecordCount >= 1 Then
                    If ValidateData("Item_Code", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_DrgNo = '" & RsObjInsert.Fields("cust_drgno").Value & "' And Item_Code = '" & strItemCode & "'") Then
                        strsql = "UPDATE DailyMKTSchedule Set Schedule_quantity=Schedule_quantity-" & Val(Rs.Fields("Quantity").Value) & ", Upd_UserId = 'DI QUERY', Upd_dt = getdate() Where Unit_Code='" + gstrUNITID + "' AND  Status = 1 and"
                        strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "'"
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    '''         If Len(Find_Value("select product_no from forecast_mst WHERE customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & Format(RsObjInsert.Fields("sch_date"), "dd mmm yyyy") & "' and Enagare_Unloc='" & RsObjInsert.Fields("unloc") & "'")) > 0 Then
                    '''             mP_Connection.Execute "Update forecast_mst Set Quantity=Quantity -" & val(Rs.Fields("Quantity")) & ",Upd_dt=Getdate(),upd_userid='DI SPARES'  WHERE customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & Format(RsObjInsert.Fields("sch_date"), "dd mmm yyyy") & "' and Enagare_UNLOC='" & RsObjInsert.Fields("unloc") & "'"
                    '''         End If
                    mP_Connection.Execute("DELETE FROM mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' AND kanbanno='" & Trim(RsObjInsert.Fields("kanbanno").Value) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                ''----Insert the record containing new KanbanNo into the table mkt_enagaredtl
                'Changed for Issue ID eMpro-20090424-30570 Starts -Insert Po type from cust_ord_hdr
                mP_Connection.Execute("Insert Into mkt_enagaredtl(Account_code,Item_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,scheduletype,UNIT_CODE,price_change_flag,batch_code,price) VALUES ( '" & stracccode & "' ,'" & strItemCode & "','" & RsObjInsert.Fields("cust_drgno").Value & "','" & RsObjInsert.Fields("quantity").Value & "','" & RsObjInsert.Fields("unloc").Value & "','" & RsObjInsert.Fields("usloc").Value & "','" & RsObjInsert.Fields("kanbanno").Value & "','" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "','" & RsObjInsert.Fields("sch_time").Value & "','" & strPOType & "','" + gstrUNITID + "','" & RsObjInsert("price_change_flag").Value & "','" & RsObjInsert("batch_code").Value & "','" & RsObjInsert("price").Value & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                'Changed for Issue ID eMpro-20090424-30570 Ends
                RsObjInsert.MoveNext()
            End While
        Else
            mP_Connection.RollbackTrans()
            MsgBox("No data Found for insertion", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        End If
        ''----To read data from tables cust_ord_hdr, cust_ord_dtl, customer_mst, tmp_enagarodtl to insert into DailyMKTschedule table
        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
        RsObjQuery.Open("select  A.Account_code,A.item_code,B.cust_drgno ,sum(B.quantity) as TotQty ,B.sch_date from mkt_enagaredtl A, tmp_enagarodtl B WHERE A.UNIT_CODE=B.UNIT_CODE AND  A.Unit_Code='" + gstrUNITID + "' AND A.KanbanNo = B.KanbanNo and A.Cust_Drgno = B.Cust_Drgno and B.session_ID='" & gstrIpaddressWinSck & "' group by A.Account_code,A.item_code,B.cust_drgno,B.sch_date ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjQuery.EOF Then
            While Not RsObjQuery.EOF
                dblqty = Val(RsObjQuery.Fields("TotQty").Value)
                stracccode = Trim(TxtCustCode.Text)
                strSbuItCode = RsObjQuery.Fields("cust_drgNo").Value
                strItemCode = RsObjQuery.Fields("Item_Code").Value
                intYYYYMM = CInt(VB6.Format(RsObjQuery.Fields("sch_date").Value, "yyyymm")) 'Date in format YYYYMM
                ''----To check if Record already exist in the DailyMKTSchedule or not
                If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' ") Then
                    ''----Item exist in DailyMKTSchedule so delete from MonthlyMKTSchedule
                    strsql = " Delete From MonthlyMKTSchedule Where Unit_Code='" + gstrUNITID + "' AND Account_Code = '" & stracccode & "'"
                    strsql = strsql & " And Cust_DrgNo = '" & strSbuItCode & "' AND Item_Code = '" & strItemCode & "'"
                    strsql = strsql & " AND Status = 1 AND Year_Month = " & intYYYYMM
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    ''----Read despatch from DailyMKTSchedule
                    dblDispatchqty = Val(SelectDataFromTable("Despatch_Qty", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status=1"))
                    ''----Read Schedule from DailyMKTSchedule
                    dblPrevSchedQty = Val(SelectDataFromTable("Schedule_Quantity", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status=1"))
                    ''----Read Max Revision No. from DailyMKTSchedule
                    intMaxRNo = CShort(Val(SelectDataFromTable("RevisionNo", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1")))
                    ''----Update DailyMKTSchedule by incrementing revision No. by 1 and setting status=0
                    strsql = "UPDATE DailyMKTSchedule set Status = 0, Upd_UserId = 'DI QUERY', Upd_dt = getdate() Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_COde = '" & strItemCode & "' and status = 1"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    ''----Insert new record with Revision No. = Max(RevisionNo)+1 and status=1
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE ) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & " ', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    strsql = strsql & dblqty + dblPrevSchedQty & " ," & dblDispatchqty & " ,1"
                    strsql = strsql & "," & intMaxRNo + 1 & ",getdate(),'DI QUERY',getdate(),'DI QUERY','" + gstrUNITID + "' )"
                Else ''----Entry does't exist in the DailyMKTSchedule
                    ''----Insert new record with Revision No.= 0 and status = 1
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE ) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    strsql = strsql & dblqty & " ,0 ,1"
                    strsql = strsql & ",0,getdate(),'DI QUERY',getdate(),'DI QUERY','" + gstrUNITID + "' )"
                End If
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                RsObjQuery.MoveNext()
            End While
        End If
        mP_Connection.CommitTrans()
        MsgBox("File has been uploaded successfully !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        txtDBFFilePath.Text = ""
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
        If Rs.State = ADODB.ObjectStateEnum.adStateOpen Then Rs.Close()
        RsObjInsert = Nothing
        RsObjQuery = Nothing
        Rs = Nothing
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        If Err.Number = -2147217900 Then
            mP_Connection.RollbackTrans()
            MsgBox("Data already uploaded. Quitting the job", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        End If
        If Err.Number = -2147217833 Then
            mP_Connection.RollbackTrans()
            MsgBox("Invalid Schedule Selection. Quitting the job", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub UpdateDAILYSchedule(ByRef strSchFILE As String)
        '-------------------------------------------------------------------------------------------------------------------------------------------
        ' Revision Date             -   27/03/2006
        ' Revision By               -   Davinder Singh
        ' Issue ID                  -   17378
        ' Revision History          -   1) DailyMKTSchedule table was not properly updated
        '                               2) Problem of linking of two Item_codes with same Cust_drgno of same customer solved
        '                               3) In SO Checking PO_Type 'O' was not considered in some cases
        '                               4) In case of uploading same file repeatedly Schedule_Qty. was not deducted from the Dailymktschedule
        '------------------------------------------------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim FSODIDchedules As Scripting.FileSystemObject
        Dim FSOReadStatus As Scripting.TextStream
        Dim strMasterString As String
        Dim ArrMasterArray() As String
        Dim ArrSplitData() As String
        Dim strstatus As String
        Dim i As Short
        Dim strSbuItCode As String
        Dim strFileDtTm As String
        Dim stracccode As String 'Customer Code
        Dim strsql As String 'SQL String
        Dim strSQLF As String
        Dim strItemCode As String 'Item Code
        Dim dblDispatchqty As Double 'Dispatch Qty used To Transfer Dispatch to New Revision No
        Dim dblPrevSchedQty As Double 'Previous Sched Qty , Added to New Qty
        Dim iQty As Double 'Hold the current Quantity to update
        Dim intYYYYMM As Integer
        Dim rstResultSet As New ADODB.Recordset
        Dim cnnDB As ADODB.Connection
        Dim iCtr As Short 'Holds Array element Counter
        Dim mstrinsert() As String 'Array to split Master data
        Dim mstrdata As String 'String to insert into MKT_EnagareDtl table
        Dim mstrinsdata() As String 'Split Data to insert into MKT_EnagareDtl table
        Dim intMaxRNo As Short 'Dav
        Dim RsObjInsert As New ADODB.Recordset
        Dim RsObjQuery As New ADODB.Recordset
        Dim Rs As ADODB.Recordset
        Dim dblCurrSchedQty As Double 'Previous Sched Qty , Added to New Qty
        Dim blnchangeddata As Boolean
        Dim strkanbanno As String
        Dim intloopcounter As Integer
        FSODIDchedules = New Scripting.FileSystemObject
        RsObjInsert = New ADODB.Recordset
        RsObjQuery = New ADODB.Recordset
        Rs = New ADODB.Recordset
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        Dim strIsNagareExists As Boolean = IseNagareRequired()   '---VIPIN DUBEY ON 31 AUG 2018 19:39 PM
        On Error Resume Next
        mP_Connection.Execute("Delete from tempdata where unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        On Error GoTo ErrHandler
        mP_Connection.BeginTrans()
        'strsql = "CREATE TABLE tempdata( [tmp_accountcode] char(12) NOT NULL , [tmp_custdrgno] char(40) NOT NULL, [tmp_itemcode] char(16)NOT NULL, [tmp_transdate] [datetime] NOT NULL, [tmp_qty] [decimal] NOT NULL, [tmp_entuserid] char(16) NOT NULL ) "
        'mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''----Read the data from textfile and put it into the temporary table Tmp_Enagarodtl
        FSOReadStatus = FSODIDchedules.OpenTextFile(txtDBFFilePath.Text, Scripting.IOMode.ForReading, False)
        mP_Connection.Execute("DELETE FROM Tmp_Enagarodtl Where Unit_Code='" + gstrUNITID + "' AND Session_id='" & Trim(gstrIpaddressWinSck) & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        While Not FSOReadStatus.AtEndOfLine
            strMasterString = ""
            strstatus = FSOReadStatus.ReadLine()
            ArrSplitData = Split(strstatus, " ")
            'The Following will create the string with relevant data to be inserted into the Tmp_enagaredtl Table
            For i = 0 To UBound(ArrSplitData)
                If Len(Trim(ArrSplitData(i))) > 0 Then
                    strMasterString = strMasterString & ArrSplitData(i) & "»"
                End If
            Next
            strMasterString = Mid(strMasterString, 1, Len(strMasterString) - 1)
            ArrMasterArray = Split(strMasterString, "»")
            If UBound(ArrMasterArray) = 7 Then
                If IsDate(ArrMasterArray(1)) Then
                    'the Array would always contain the data as follows :
                    '192.168.35.35 » M424     »36610M844P1»  24  »  A1-1 »  A11 »17L303057372A11»17-NOV-2003» 15:00
                    'SESSION ID » VENDOR CODE » CUSTDRGNO » QTY » UNLOC » USLOC » KANBAN NO     » SCH DATE  » SCH TIME ¬
                    mP_Connection.Execute("INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,UNIT_CODE) values('" & gstrIpaddressWinSck & "','" & Trim(ArrMasterArray(0)) & "','" & Trim(ArrMasterArray(3)) & "' ,'" & Trim(ArrMasterArray(4)) & "','" & Trim(ArrMasterArray(5)) & "','" & Trim(ArrMasterArray(6)) & "','" & Trim(ArrMasterArray(7)) & "','" & Trim(ArrMasterArray(1)) & "','" & Trim(ArrMasterArray(2)) & "','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
        End While
        'Added By Arshad - used in montlhy schedule uploading to remove nagare that is being used for planning
        'mP_Connection.Execute "insert into mkt_enagaredtl_tentative select * from mkt_enagaredtl where kanbanNo like 'Nagare%'"
        mP_Connection.Execute("insert into mkt_enagaredtl_tentative (Account_code,Item_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Freq,UNIT_CODE) select Account_code,Item_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,convert(varchar(11),Sch_date,106),Sch_time,Freq,UNIT_CODE from mkt_enagaredtl where kanbanNo like 'Nagare%' AND Unit_Code='" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute("Delete from Tmp_enagarodtl Where Unit_Code='" + gstrUNITID + "' AND kanbanNo like 'Nagare%'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''----To read the data from Tmp_enagarodtl and put it into the Recordset
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RsObjInsert.Open("SELECT * FROM Tmp_enagarodtl Where Unit_Code='" + gstrUNITID + "' AND Session_ID='" & gstrIpaddressWinSck & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'prashant 
        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
        RsObjQuery.Open("select isnull(ENAGARE_ALLOWED_ALREADYINVOICE,0)as ENAGARE_ALLOWED_ALREADYINVOICE  FROM sales_parameter Where Unit_Code='" + gstrUNITID + "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If RsObjQuery.Fields(0).Value = True Then 'Value is set for eNagareUploadingOnBasisOfSO in sales_parameter
            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
            RsObjQuery.Open("SELECT KANBANNO FROM VW_ENAGAREUPLOAD_ALREADYINVOICE Where Unit_Code='" + gstrUNITID + "' AND Session_ID='" & Trim(gstrIpaddressWinSck) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If RsObjQuery.RecordCount > 0 Then 'There is no SO Active & Authorized
                For intloopcounter = 1 To RsObjQuery.RecordCount
                    If strkanbanno = "" Then
                        strkanbanno = RsObjQuery.Fields("KANBANNO").Value.ToString
                    Else
                        strkanbanno = strkanbanno & "," & RsObjQuery.Fields("KANBANNO").Value.ToString
                    End If
                Next
                MsgBox(" The Invoice for Nagare Nos. " & strkanbanno & "  has already been generated. File cannot be upload  for selected Customer. " & vbCrLf & " It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                mP_Connection.RollbackTrans()
                ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                Exit Sub
            End If
        End If
        '
        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                'To retrieve Customer code line by line
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                stracccode = ""
                RsObjQuery.Open("SELECT customer_code FROM customer_mst Where Unit_Code='" + gstrUNITID + "' And cust_vendor_code='" & Trim(RsObjInsert.Fields("vendor_code").Value) & "' and customer_code = '" & Trim(Me.TxtCustCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjQuery.EOF Then
                    stracccode = Trim(RsObjQuery.Fields("customer_code").Value)
                Else
                    MsgBox("No Data found in the Customer Master for the combination of seleted Customer Code[" & Trim(TxtCustCode.Text) & "] and customer vendor code[" & Trim(RsObjInsert.Fields("vendor_code").Value) & "] in the file.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Sub
                End If
                ''----Read the item_code from the custitem_mst table
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                strItemCode = ""
                RsObjQuery.Open("SELECT item_code FROM custitem_mst Where Unit_Code='" + gstrUNITID + "' And cust_drgno='" & Trim(RsObjInsert.Fields("Cust_drgno").Value) & "' AND account_code='" & stracccode & "' and active=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjQuery.EOF Then
                    strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                End If
                'Changed for More than one item code active for more
                'than one SO authorized and active then pick depending on Sales_parameter
                If RsObjQuery.RecordCount > 1 Then
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("SELECT D.Item_code from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE=D.UNIT_CODE  AND H.Unit_Code='" + gstrUNITID + "' AND H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type='O' and D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If Not RsObjQuery.EOF Then
                        strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                        GoTo Onerec
                    Else
                        If MsgBox(" There are more than 1 item code defined for this Customer part Code : " & Trim(RsObjInsert.Fields("cust_drgno").Value) & "." & vbCrLf & " Proceed with it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                            GoTo Onerec
                        Else
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Sub
                        End If
                    End If
                ElseIf RsObjQuery.RecordCount < 1 Then  'Message for Item code is not Active and roll back the uploading
                    MsgBox(" Item Code not found for Cust Part Code code : " & Trim(RsObjInsert.Fields("Cust_drgno").Value) & vbCrLf & " Please correct the data first. It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else 'There is only one row / Item Code for the defined Internal Code
                    'Code Added by Arshad on 05/04/2005
Onerec:
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("select eNagareUploadingOnBasisOfSO FROM sales_parameter Where Unit_Code='" + gstrUNITID + "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If RsObjQuery.Fields(0).Value = True Then 'Value is set for eNagareUploadingOnBasisOfSO in sales_parameter
                        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                        RsObjQuery.Open("SELECT D.Item_code from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE=D.UNIT_CODE AND  H.Unit_Code='" + gstrUNITID + "' AND H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type='O' and D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If RsObjQuery.RecordCount = 0 Then 'There is no SO Active & Authorized
                            MsgBox(" There is no SO Authorized and Active for Item " & Trim(RsObjInsert.Fields("Cust_Drgno").Value) & " for selected Customer. " & vbCrLf & " It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Sub
                        Else 'Item code is Fetched to be inserted into the table MKT_EnagareDtl
                            strItemCode = Trim(RsObjQuery.Fields("Item_Code").Value)
                        End If
                    End If
                End If
                '----case specifically added for sun vacuum--if same nagarro nos is repeated,then
                'delete the previous nos from mktenagaro_dtl and insert new one .
                If ChkNagarroNo(RsObjInsert.Fields("kanbanno").Value) Then
                    blnchangeddata = False
                    If ValidateData("Item_Code", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_DrgNo = '" & Trim(RsObjInsert.Fields("Cust_drgno").Value) & "' and status=1 And Item_Code = '" & strItemCode & "'") Then
                        dblPrevSchedQty = CDbl(SelectDataFromTable("Quantity", "mkt_enagaredtl", " Kanbanno='" & RsObjInsert.Fields("kanbanno").Value & "'"))
                        'prashant rajpal changed as on 10 june 2009
                        dblCurrSchedQty = CDbl(SelectDataFromTable("Quantity", "tmp_enagarodtl", " Kanbanno='" & RsObjInsert.Fields("kanbanno").Value & "'"))
                        'prashant rajpal changed ended as on 10 june 2009
                        'prashant changed as RFC
                        If schkanbanunitFlag = True Then
                            If CDbl(dblPrevSchedQty) < CDbl(dblCurrSchedQty) Then
                                blnchangeddata = True
                                strsql = " insert into tempdata select account_code,cust_drgno,item_code,trans_date,schedule_quantity,ent_userid,UNIT_CODE  from dailymktschedule where "
                                '10263152 
                                'strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' and unit_code='" & gstrDateFormat & "'"
                                strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' and unit_code='" & gstrUNITID & "'"
                                '10263152 
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'UPDATE REVISION NOs
                                'prashant changed ended as RFC
                                'prashant rajpal changed as on 10 june 2009
                                'strsql = "UPDATE DailyMKTSchedule Set Schedule_quantity=Schedule_quantity-" & dblPrevSchedQty & ", Upd_Userid='E Nagare', Upd_dt=getdate() where  Status = 1 and"
                                'strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "'"
                                strsql = "UPDATE DailyMKTSchedule Set Schedule_flag=0 ,status=0, Upd_Userid='E Nagare', Upd_dt=getdate() Where Unit_Code='" + gstrUNITID + "' AND  Status = 1 and "
                                strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "'"
                                'prashant rajpal changed ended as on 10 june 2009
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'UPDATE REVISION NOs
                            End If
                        Else
                            If CDbl(dblPrevSchedQty) <= CDbl(dblCurrSchedQty) Then
                                blnchangeddata = True
                                strsql = " insert into tempdata select account_code,cust_drgno,item_code,trans_date,schedule_quantity,ent_userid,UNIT_CODE  from dailymktschedule where "
                                '10263152 
                                'strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' and unit_code='" & gstrDateFormat & "'"
                                strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' and unit_code='" & gstrUNITID & "'"
                                '10263152 
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'UPDATE REVISION NOs
                                'prashant changed ended as RFC
                                'prashant rajpal changed as on 10 june 2009
                                'strsql = "UPDATE DailyMKTSchedule Set Schedule_quantity=Schedule_quantity-" & dblPrevSchedQty & ", Upd_Userid='E Nagare', Upd_dt=getdate() where  Status = 1 and"
                                'strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "'"
                                strsql = "UPDATE DailyMKTSchedule Set Schedule_flag=0 ,status=0, Upd_Userid='E Nagare', Upd_dt=getdate() Where Unit_Code='" + gstrUNITID + "' AND  Status = 1 and "
                                strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "'"
                                'prashant rajpal changed ended as on 10 june 2009
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'UPDATE REVISION NOs
                            End If
                        End If
                        
                    Else
                        blnchangeddata = True
                    End If
                    If blnchangeddata = True Then
                        If strIsNagareExists Then
                            mP_Connection.Execute("UPDATE mkt_enagaredtl SET Quantity='" & RsObjInsert.Fields("quantity").Value & "' WHERE Unit_Code='" + gstrUNITID + "' AND kanbanno='" & RsObjInsert.Fields("kanbanno").Value & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        Else
                            mP_Connection.Execute("DELETE FROM mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' AND kanbanno='" & RsObjInsert.Fields("kanbanno").Value & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute("Insert Into mkt_enagaredtl(Account_code,Item_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Scheduletype,UNIT_CODE) VALUES ( '" & stracccode & "' ,'" & strItemCode & "','" & RsObjInsert.Fields("cust_drgno").Value & "','" & RsObjInsert.Fields("quantity").Value & "','" & RsObjInsert.Fields("unloc").Value & "','" & RsObjInsert.Fields("usloc").Value & "','" & RsObjInsert.Fields("kanbanno").Value & "','" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "','" & RsObjInsert.Fields("sch_time").Value & "','O','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    Else
                        mP_Connection.Execute("DELETE FROM tmp_enagarodtl Where Unit_Code='" + gstrUNITID + "' AND kanbanno='" & RsObjInsert.Fields("kanbanno").Value & "' and session_id='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    Else  'nagare no not exists 
                        mP_Connection.Execute("Insert Into mkt_enagaredtl(Account_code,Item_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Scheduletype,UNIT_CODE) VALUES ( '" & stracccode & "' ,'" & strItemCode & "','" & RsObjInsert.Fields("cust_drgno").Value & "','" & RsObjInsert.Fields("quantity").Value & "','" & RsObjInsert.Fields("unloc").Value & "','" & RsObjInsert.Fields("usloc").Value & "','" & RsObjInsert.Fields("kanbanno").Value & "','" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "','" & RsObjInsert.Fields("sch_time").Value & "','O','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If 'nagare no loop closed
                    RsObjInsert.MoveNext()
            End While
        Else
            mP_Connection.RollbackTrans()
            MsgBox("No data Found for insertion", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.Open("select  A.Account_code,A.item_code,B.cust_drgno ,sum(B.quantity) as TotQty ,B.sch_date from mkt_enagaredtl A, tmp_enagarodtl B where A.UNIT_CODE=B.UNIT_CODE AND A.Unit_Code='" + gstrUNITID + "' AND  A.KanbanNo = B.KanbanNo and A.Cust_Drgno = B.Cust_Drgno and B.session_ID='" & gstrIpaddressWinSck & "' group by A.Account_code,A.item_code,B.cust_drgno,B.sch_date ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                iQty = RsObjInsert.Fields("TotQty").Value 'Cumulative quantity to be inserted to daily marketting schedule of that item
                stracccode = Trim(RsObjInsert.Fields("Account_code").Value) 'Account Code
                strSbuItCode = Trim(RsObjInsert.Fields("cust_drgno").Value) 'Cust_drgNo
                strItemCode = Trim(RsObjInsert.Fields("Item_Code").Value) 'Item code
                intYYYYMM = CInt(VB6.Format(RsObjInsert.Fields("sch_date").Value, "yyyymm")) 'Date in format YYYYMM
                'strkanbanno = Trim(RsObjInsert.Fields("kanbanno").Value) 'Item code
                'prashant rajpal changed as per RFC
                If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and upper(ent_userid) In ('TENTATIVE')") Then
                    strsql = "UPDATE DailyMKTSchedule_history Set RevisionNo=RevisionNo +1  Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                    strsql = strsql & " and upper(ent_userid)='TENTATIVE'"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strsql = "Insert Into DailyMKTSchedule_history "
                    strsql = strsql & " select tmp_accountcode ,tmp_transdate,tmp_itemcode,tmp_custdrgno,tmp_qty,"
                    strsql = strsql & "  0, getdate() ,tmp_entuserid,getdate(),tmp_entuserid,UNIT_CODE from tempdata Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " tmp_accountcode = '" & stracccode & "' AND tmp_transdate = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND tmp_itemcode = '" & strItemCode & "' AND tmp_custdrgno='" & strSbuItCode & "' "
                    strsql = strsql & " AND UPPER(tmp_entuserid)='TENTATIVE' "
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strsql = "delete from DailyMKTSchedule Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                    strsql = strsql & " and ent_userid in('TENTATIVE') "
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and upper(ent_userid) In ('NAGARE COMP')") Then
                    strsql = "UPDATE DailyMKTSchedule_history Set RevisionNo=RevisionNo +1  Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                    strsql = strsql & " and upper(ent_userid)='NAGARE COMP'"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strsql = "Insert Into DailyMKTSchedule_history "
                    strsql = strsql & " select tmp_accountcode ,tmp_transdate,tmp_itemcode,tmp_custdrgno,tmp_qty,"
                    strsql = strsql & "  0, getdate() ,tmp_entuserid,getdate(),tmp_entuserid,UNIT_CODE from tempdata Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " tmp_accountcode = '" & stracccode & "' AND tmp_transdate = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND tmp_itemcode = '" & strItemCode & "' AND tmp_custdrgno='" & strSbuItCode & "' "
                    strsql = strsql & " AND UPPER(tmp_entuserid)='NAGARE COMP' "
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strsql = "delete from DailyMKTSchedule Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                    strsql = strsql & " and ent_userid in('NAGARE COMP') "
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                'prashant rajpal changed ended as per RFC
                If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and status = 1  ") Then
                    'If Item HAS Monthly Schedule then delete it
                    strsql = " Delete From MonthlyMKTSchedule Where Unit_Code='" + gstrUNITID + "' AND Account_Code = '" & stracccode & "'"
                    strsql = strsql & " And Cust_DrgNo = '" & strSbuItCode & "' AND Item_Code = '" & strItemCode & "'"
                    strsql = strsql & " AND Status = 1 AND Year_Month = " & intYYYYMM
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    'If Entry for this Item is already there .... for strFields(3) Date
                    dblDispatchqty = CDbl(Val(SelectDataFromTable("Despatch_Qty", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1")))
                    dblPrevSchedQty = CDbl(Val(SelectDataFromTable("Schedule_Quantity", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1")))
                    intMaxRNo = CShort(Val(SelectDataFromTable("RevisionNo", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1")))
                    '1.)Update and INCREASE by 1 Revision No,Status = 0 for Previous Entries
                    strsql = "UPDATE DailyMKTSchedule Set Status = 0, schedule_flag=0 ,Upd_Userid='E Nagare',upd_dt=getdate() Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_COde = '" & strItemCode & "' and status = 1"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE ) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & " ', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    strsql = strsql & dblPrevSchedQty + iQty & " ," & dblDispatchqty & " ,1"
                    strsql = strsql & "," & intMaxRNo + 1 & ",getdate(),'E Nagare',getdate(),'E Nagare','" + gstrUNITID + "' )"
                Else
                    'Insert Item into DailyMKTSchedule
                    dblDispatchqty = CDbl(Val(SelectDataFromTable("Despatch_Qty", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "'")))
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, " dd mmm yyyy") & "', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    'strsql = strsql & CDbl(iQty) & " ,0 ,1"
                    strsql = strsql & CDbl(iQty) & " ," & dblDispatchqty & ",1"
                    strsql = strsql & ",0,getdate(),'E Nagare',getdate(),'E Nagare','" + gstrUNITID + "')"
                End If
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                RsObjInsert.MoveNext()
            End While
        End If
        mP_Connection.CommitTrans()
        '---- PiCKLIST GENERATION FOR SANAND UNIT
        If IseNagareRequired() Then

            Dim sqlcmd As New System.Data.SqlClient.SqlCommand
            sqlcmd.Connection = SqlConnectionclass.GetConnection
            sqlcmd.CommandTimeout = 0
            sqlcmd.CommandType = CommandType.StoredProcedure
            sqlcmd.Parameters.Clear()

            If gstrUnitId.ToUpper() = "MTM" Or gstrUnitId.ToUpper() = "MR1" Or gstrUnitId.ToUpper() = "MYT" Then
                sqlcmd.CommandText = "USP_GENERATE_ENAGARE_PICKLIST_MTM"
            Else
                sqlcmd.CommandText = "USP_GENERATE_ENAGARE_PICKLIST"
            End If
            sqlcmd.Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
            sqlcmd.Parameters.Add("@ACCOUNT_CODE", SqlDbType.VarChar, 100).Value = TxtCustCode.Text.ToString.Trim
            sqlcmd.Parameters.Add("@USER_ID", SqlDbType.VarChar, 50).Value = mP_User
            sqlcmd.ExecuteNonQuery()
        End If

        MsgBox("File has been uploaded successfully !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        txtDBFFilePath.Text = ""
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        If Err.Number = -2147217833 Then
            mP_Connection.RollbackTrans()
            MsgBox("Invalid Schedule Selection." & vbCrLf & "Please Select Correct Schedule Option.", MsgBoxStyle.Information, ResolveResString(100))
            txtDBFFilePath.ForeColor = System.Drawing.Color.Red
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        If Err.Number = -2147217900 Then
            mP_Connection.RollbackTrans()
            MsgBox("Data already uploaded", MsgBoxStyle.Information, ResolveResString(100))
            txtDBFFilePath.ForeColor = System.Drawing.Color.Red
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0029_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mlngFormTag
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0029_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0029_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '-----------------------------------------------------------------------
        'Escape Key Handling
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be unloaded
                If MsgBox("Want To Close This Screen ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "eMPro") = MsgBoxResult.Yes Then
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
    Private Sub frmMKTTRN0029_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.AppStarting)
        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlUploadSchedulesHDR.HeaderString)
        Call FitToClient(Me, fraMain, ctlUploadSchedulesHDR, lblUploadCmd, 300)
        'Setting Print and Close Buttons
        cmdTransfer.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lblUploadCmd.Left) + 70)
        cmdTransfer.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblUploadCmd.Top) + 50)
        cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdTransfer.Left) + VB6.PixelsToTwipsX(cmdTransfer.Width) + 10)
        cmdClose.Top = cmdTransfer.Top
        OptNagare.Checked = True
        CmdCustCode.Enabled = True
        'Added for Issue ID eMpro-20090505-31005 Starts
        lblSOtype.Visible = False
        cmbsotype.Visible = False
        Call ChkVBSchUpdFlag()
        Call ChkKanbannunit()
        'Added for Issue ID eMpro-20090505-31005 Ends
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0029_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        'Assign form to nothing
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '------------------------------------------------------------------------------
        'Created By     -   Nitin Sood
        'Description    -   Get Data from BackEnd
        '------------------------------------------------------------------------------
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & " Where Unit_Code='" + gstrUNITID + "' AND " & mstrCondition
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
        strsql = "Select " & mstrFieldName & " From " & mstrTableName & " Where Unit_Code='" + gstrUNITID + "' AND " & mstrCond
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
    Private Sub ShowCode_Desc(ByVal pstrQuery As String, ByRef pctlCode As System.Windows.Forms.TextBox, Optional ByRef pctlDesc As System.Windows.Forms.Label = Nothing)
        '--------------------------------------------------------------------------------------
        'Name       :   ShowCode_Desc
        'Type       :   Sub
        'Author     :   Jasmeet Singh Bawa
        'Arguments  :   Query(string),Code(Text Box),Description(Label)
        'Return     :   None
        'Purpose    :   Show Code and Description window and set focus on code
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        'Changing the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        mstrHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, pstrQuery, "Code Help", 2)
        'Changing the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(mstrHelp) <> -1 Then
            If mstrHelp(0) <> "0" Then
                pctlCode.Text = Trim(mstrHelp(0))
                If Not (pctlDesc Is Nothing) Then
                    pctlDesc.Text = Trim(mstrHelp(1))
                End If
                If pctlCode.Enabled Then pctlCode.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
        'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub OptDI_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptDI.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            TxtCustCode.Text = ""
            TxtCustCode.Enabled = True
            TxtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            CmdCustCode.Enabled = True
            LblCustDesc.Text = ""
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub OptDI_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptDI.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub OptDISpares_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptDISpares.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            TxtCustCode.Text = ""
            TxtCustCode.Enabled = True
            TxtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            CmdCustCode.Enabled = True
            LblCustDesc.Text = ""
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub OptDISpares_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptDISpares.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub optNagare_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptNagare.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            TxtCustCode.Text = ""
            TxtCustCode.Enabled = True
            TxtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            CmdCustCode.Enabled = True
            LblCustDesc.Text = ""
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub OptNagare_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptNagare.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCustCode.TextChanged
        On Error GoTo ErrHandler
        If Len(TxtCustCode.Text) = 0 Then
            LblCustDesc.Text = ""
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCustCode.Enter
        On Error GoTo ErrHandler
        With TxtCustCode
            .SelectionStart = 0
            .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtCustCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdcustcode_Click(CmdCustCode, New System.EventArgs())
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtcustcode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles TxtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim RsObjCustCodeVal As New ADODB.Recordset
        On Error GoTo ErrHandler
        If RsObjCustCodeVal.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjCustCodeVal.Close()
        If Len(TxtCustCode.Text) > 0 Then
            'Changes against 10737738 
            If SchUpdFlag = True Then
                RsObjCustCodeVal.Open("SELECT customer_code,Cust_name FROM customer_mst Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & Trim(TxtCustCode.Text) & "' and SCH_UPLOAD_CODE='NAGARE' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Else
                RsObjCustCodeVal.Open("SELECT customer_code,Cust_name FROM customer_mst Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & Trim(TxtCustCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            End If
            If Not RsObjCustCodeVal.EOF Then
                TxtCustCode.Text = Trim(RsObjCustCodeVal.Fields(0).Value)
                LblCustDesc.Text = Trim(RsObjCustCodeVal.Fields(1).Value)
            Else
                MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                TxtCustCode.Text = ""
                LblCustDesc.Text = ""
                Cancel = True
                GoTo EventExitSub
            End If
        End If
            GoTo EventExitSub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
            eventArgs.Cancel = Cancel
    End Sub
    Private Function DISchedules() As Object
        '-------------------------------------------------------------------------------------------------------------------------------------------
        ' Revision Date             -   27/03/2006
        ' Revision By               -   Davinder Singh
        ' Issue ID                  -   17378
        ' Revision History          -   Function changed to send data only in the Forecast_Mst Table
        '                               If KanbanNo is missing then send 'TENTITIVE' in the field
        '------------------------------------------------------------------------------------------------------------------------------------------
        Dim FSODIDchedules As New Scripting.FileSystemObject
        Dim FSOReadStatus As Scripting.TextStream
        Dim strstatus As String
        Dim i As Short
        Dim dblqty As Double
        Dim strMasterString As String
        Dim ArrMasterArray() As String
        Dim ArrSplitData() As String
        Dim RsObjInsert As New ADODB.Recordset
        Dim RsObjQuery As New ADODB.Recordset
        Dim RsObjItemcode As New ADODB.Recordset
        Dim RsObjCUSTDRG As New ADODB.Recordset
        Dim stracccode As String
        Dim strItemCode As String
        Dim strSbuItCode As String
        Dim intYYYYMM As Integer
        Dim strsql As String
        Dim dblDispatchqty As Double
        Dim dblPrevSchedQty As Double
        Dim strunLoc As String ' Declare by Sandeep
        Dim Intcounter As Short
        Dim strfirstdate As String
        Dim iQty As Double
        Dim intMaxRNo As Short
        Dim CreateString As String
        Dim StrSQLQuery As String
        Dim strKanbanNo As String
        On Error Resume Next
        mP_Connection.Execute("DELETE FROM table tempitemcode WHERE UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        On Error GoTo ErrHandler
        'strsql = "CREATE TABLE tempitemcode( [tmp_accountcode] char(12) NOT NULL , [tmp_custdrgno] char(40) NOT NULL, [tmp_itemcode] char(16)NOT NULL, [tmp_transdate] [datetime] NOT NULL)"
        'mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.BeginTrans()
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        FSOReadStatus = FSODIDchedules.OpenTextFile(txtDBFFilePath.Text, Scripting.IOMode.ForReading, False)
        mP_Connection.Execute("DELETE FROM Tmp_Enagarodtl Where Unit_Code='" + gstrUNITID + "' AND Session_id='" & Trim(gstrIpaddressWinSck) & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Intcounter = 1
        While Not FSOReadStatus.AtEndOfLine
            strMasterString = ""
            strstatus = FSOReadStatus.ReadLine()
            ArrSplitData = Split(strstatus, " ")
            'The Following will create the string with relevant data to be inserted into the Tmp_enagaredtl Table
            For i = 0 To UBound(ArrSplitData)
                If Len(Trim(ArrSplitData(i))) > 0 Then
                    strMasterString = strMasterString & ArrSplitData(i) & "»"
                End If
            Next
            strMasterString = Mid(strMasterString, 1, Len(strMasterString) - 1)
            'String will be delivered as follows in following formats :
            'S200»15-APR-2004»39231M70F00-T01»228»A2-4»NAGARE»COMP»1-1»75»
            'S200»15-APR-2004»43250M77500»94»A1-6»D0448043176»1-1»100
            'S200»29-APR-2004»11309M72F00»6»E2-1
            'Split the string to insert data into Table Tmp_enagaredtl
            ArrMasterArray = Split(strMasterString, "»")
            'Debug.Print UBound(ArrMasterArray)
            'String data will contain data with DI No. / NAGARE COMP / Tentative Data
            'String will be handled accordingly to insert data
            If UBound(ArrMasterArray) = 7 Then
                If IsDate(ArrMasterArray(1)) Then
                    If Intcounter = 1 Then
                        strfirstdate = ArrMasterArray(1)
                    End If
                    If Len(ArrMasterArray(2)) = 5 Then
                        mP_Connection.RollbackTrans()
                        MsgBox("Invalid Schedule option Selected. File is E Nagare Schedule.", MsgBoxStyle.Information, ResolveResString(100))
                        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        Exit Function
                    Else
                        'prashant rajpal changed as on 05 feb 2008
                        If VB6.Format(CDate(ArrMasterArray(1)), "yyyymm") < VB6.Format(GetServerDate(), "YYYYMM") Then
                            MsgBox("Schedule can not be uploaded for previous month ", MsgBoxStyle.OkOnly, ResolveResString(100))
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        End If
                        'if ArrMasterArray(1)
                        'NAGARE and COMP are considered as two seperate words .So they are concatenated and inserted into the table
                        mP_Connection.Execute("INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Freq,UNIT_CODE) values('" & gstrIpaddressWinSck & "','" & Trim(ArrMasterArray(0)) & "','" & Trim(ArrMasterArray(2)) & "' ,'" & Trim(ArrMasterArray(3)) & "','" & Trim(ArrMasterArray(4)) & "',' ','" & Trim(ArrMasterArray(5)) & " " & Trim(ArrMasterArray(6)) & "','" & Trim(ArrMasterArray(1)) & "','11:59','" & ArrMasterArray(7) & "','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        Intcounter = Intcounter + 1
                    End If
                End If
            End If
            If UBound(ArrMasterArray) = 5 Then 'DI/KANBAN is missing :Tentative Schedule
                If IsDate(ArrMasterArray(1)) Then 'DI No is missing : it means it is a Tentative Schedule :: so it will be pushed into Forecast Master
                    If Len(ArrMasterArray(2)) = 5 Then
                        mP_Connection.RollbackTrans()
                        MsgBox("Invalid Schedule option Selected. File is E Nagare Schedule.", MsgBoxStyle.Information, ResolveResString(100))
                        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        Exit Function
                    Else
                        mP_Connection.Execute("INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,UNIT_CODE) values('" & Trim(gstrIpaddressWinSck) & "','" & Trim(ArrMasterArray(0)) & "','" & Trim(ArrMasterArray(2)) & "' ,'" & Trim(ArrMasterArray(3)) & "','" & Trim(ArrMasterArray(4)) & "',' ','TENTATIVE','" & Trim(ArrMasterArray(1)) & "','11:59','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
            End If
        End While
        FSODIDchedules = Nothing
        FSOReadStatus = Nothing
        'prashant rajpal changed as on 24 jan 2008 for issue id: 22228
        strsql = " Select * from  forecast_mst Where Unit_Code='" + gstrUNITID + "' AND   customer_code =  '" & TxtCustCode.Text & "' and due_date >= '" & strfirstdate & "' and   Enagare_UNLOC not in('N/A','FCST') "
        gobjDB.GetResult(strsql)
        mP_Connection.Execute("Set dateFormat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        gobjDB.MoveFirst()
        'for update revision no  ie revision no=revision no +1
        If gobjDB.RowCount >= 1 Then
            mP_Connection.Execute("UPDATE Forecast_Mst_History Set RevisionNo = RevisionNo + 1 Where Unit_Code='" + gstrUNITID + "' AND  Customer_Code = '" & TxtCustCode.Text & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        While Not gobjDB.EOFRecord
            strItemCode = ""
            'RsObjCUSTDRG.Open "SELECT item_code FROM custitem_mst WHERE cust_drgno='" & strSbuItCode & "' AND account_code='" & stracccode & "'", mP_Connection, adOpenStatic, adLockReadOnly
            'If Not RsObjCUSTDRG.EOF Then
            '   stritemcode = Trim(RsObjCUSTDRG.Fields("Item_code")) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
            'End If
            mP_Connection.Execute("insert into forecast_mst_history (Customer_code,Product_no,Due_date,Quantity,RevisionNo,ScheduleNo,ent_dt,ent_userid,upd_dt,upd_userid,Enagare_UNLOC,UNIT_CODE )   Values ('" & TxtCustCode.Text & "' ,'" & gobjDB.GetValue("product_no") & "', '" & gobjDB.GetValue("due_date") & "', '" & gobjDB.GetValue("Quantity") & "', '0',  '" & 1 & "','" & gobjDB.GetValue("ent_dt") & "', '" & gobjDB.GetValue("ent_userid") & "', '" & gobjDB.GetValue("upd_dt") & "' ,'" & gobjDB.GetValue("upd_userid") & "','" & gobjDB.GetValue("Enagare_UNLOC") & "','" + gstrUNITID + "')  ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            gobjDB.MoveNext()
        End While
        'prashant rajpal changed as on 13 nov 2008 for issue id:-eMpro-20081113-23543
        'mP_Connection.Execute "delete From  forecast_mst Where customer_code =  '" & txtCustCode.Text & "' and due_date >= '" & strfirstdate & "' and   Enagare_UNLOC <> 'N/A' "
        mP_Connection.Execute("delete From forecast_mst Where Unit_Code='" + gstrUNITID + "' AND customer_code =  '" & TxtCustCode.Text & "' and due_date >= '" & strfirstdate & "' and   Enagare_UNLOC not in('N/A','FCST') ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        'prashant rajpal changed ended as on 13 nov 2008 for issue id:-eMpro-20081113-23543
        'mP_Connection.Execute "delete From  forecast_mst Where customer_code =  '" & TxtCustCode.Text & "' and due_date >= '" & strMstFields(5) & "'  and   Enagare_UNLOC <>  'N/A' "
        'prashant rajpal changed ended
        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
        RsObjQuery.Open("Select Vendor_code,Cust_drgNo,Sum(Quantity) as TotQty,UNLOC,KanBanno,Sch_Date from Tmp_Enagarodtl Where Unit_Code='" + gstrUNITID + "' AND session_id = '" & gstrIpaddressWinSck & "'  group by Vendor_code,Cust_drgNo,UNLOC,KanBanno,Sch_Date order by KanBanno desc ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjQuery.EOF Then
            While Not RsObjQuery.EOF
                dblqty = RsObjQuery.Fields("TotQty").Value 'Cumulative quantity to be inserted to daily marketting schedule of that item
                strSbuItCode = RsObjQuery.Fields("cust_drgNo").Value 'Cust_drgNo
                strunLoc = RsObjQuery.Fields("UNLoc").Value ' UNLOC
                If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
                stracccode = ""
                RsObjInsert.Open("SELECT customer_code FROM customer_mst Where Unit_Code='" + gstrUNITID + "' AND cust_vendor_code='" & Trim(RsObjQuery.Fields("vendor_code").Value) & "' and customer_code = '" & Trim(Me.TxtCustCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjInsert.EOF Then
                    stracccode = Trim(RsObjInsert.Fields("customer_code").Value)
                Else
                    MsgBox("No Data found in the Customer Master for the combination of seleted Customer Code[" & Trim(TxtCustCode.Text) & "] and customer vendor code[" & Trim(RsObjQuery.Fields("vendor_code").Value) & "] available in the file.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                ''----Read the item_code from the custitem_mst table
                If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
                strItemCode = ""
                RsObjInsert.Open("SELECT item_code FROM custitem_mst Where Unit_Code='" + gstrUNITID + "' AND cust_drgno='" & strSbuItCode & "' AND account_code='" & stracccode & "' and active=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjInsert.EOF Then
                    strItemCode = Trim(RsObjInsert.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                End If
                'Changed for More than one item code active for more
                'than one SO authorized and active then pick depending on Sales_parameter
                If RsObjInsert.RecordCount > 1 Then
                    If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
                    RsObjInsert.Open("SELECT D.Item_code from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE=D.UNIT_CODE AND H.Unit_Code='" + gstrUNITID + "' AND H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type='O' and D.Active_Flag='A' and D.cust_drgNo='" & strSbuItCode & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If Not RsObjInsert.EOF Then
                        strItemCode = Trim(RsObjInsert.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                        GoTo Onerec
                    Else
                        If MsgBox(" There are more than 1 item code defined for this Customer part Code : " & strSbuItCode & "." & vbCrLf & " Proceed with it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                            GoTo Onerec
                        Else
                            If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        End If
                    End If
                ElseIf RsObjInsert.RecordCount < 1 Then  'Message for Item code is not Active and roll back the uploading
                    MsgBox(" Item Code not found for Cust Part code : " & strSbuItCode & vbCrLf & " Please correct the data first. It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                Else 'There is only one row / Item Code for the defined Internal Code
Onerec:
                    If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
                    If Trim(UCase(RsObjQuery.Fields("kanbanno").Value)) = "NAGARE COMP" Then
                        If Len(Find_Value("select product_no from forecast_mst Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' and Enagare_Unloc='" & Trim(strunLoc) & "'")) > 0 Then
                            'strSql = "Update forecast_mst Set Quantity=" & dblqty & ",Upd_dt=Getdate(),upd_userid='NAGARE COMP' WHERE customer_code='" & stracccode & "' and product_no='" & stritemcode & "' and due_date='" & Format(RsObjQuery.Fields("sch_date"), "dd mmm yyyy") & "' and Enagare_UNLOC='" & Trim(strunLoc) & "'"
                            strsql = ""
                        Else
                            'strSql = "INSERT INTO forecast_mst(Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt, ENagare_UNLOC) VALUES ('" & stracccode & "','" & stritemcode & "' ,'" & Format(RsObjQuery.Fields("sch_date"), "dd mmm yyyy") & "'," & dblqty & ",'NAGARE COMP',getdate(),'NAGARE COMP',getdate(), '" & Trim(strunLoc) & "')"
                            strsql = ""
                        End If
                    ElseIf (Trim(UCase(RsObjQuery.Fields("kanbanno").Value)) = "TENTATIVE") Or (Mid(Trim(UCase(RsObjQuery.Fields("kanbanno").Value)), 1, 6) = "NAGARE") Then
                        If Len(Find_Value("select product_no from forecast_mst Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' and Enagare_UNloc='" & Trim(strunLoc) & "'")) > 0 Then
                            If Trim(UCase(RsObjQuery.Fields("kanbanno").Value)) = "TENTATIVE" Then
                                strsql = "Update forecast_mst Set Quantity=" & dblqty & ",upd_userid='TENTATIVE',upd_dt=Getdate() Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' and Enagare_UnLoc='" & Trim(strunLoc) & "'"
                            End If
                        Else
                            If Trim(UCase(RsObjQuery.Fields("kanbanno").Value)) = "TENTATIVE" Then
                                strsql = "INSERT INTO forecast_mst(Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt, Enagare_UNLOC,UNIT_CODE) VALUES ('" & stracccode & "','" & strItemCode & "' ,'" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "'," & dblqty & ",'TENTATIVE',getdate(),'TENTATIVE',getdate(), '" & Trim(strunLoc) & "','" + gstrUNITID + "')"
                            End If
                        End If
                    End If
                    If strsql <> "" Then
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
                RsObjQuery.MoveNext()
            End While
        Else
            mP_Connection.RollbackTrans()
            MsgBox("No Data Found to Upload", MsgBoxStyle.Information, ResolveResString(100))
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        'prashant rajpal changed for RFC requirement
        ' Delete the tentative rows if Nagare Comp Exists for the same date with the same customer
        mP_Connection.Execute(" delete Tmp_Enagarodtl  from Tmp_Enagarodtl a  Where ( select count(*) from dailymktschedule b Where B.UNIT_CODE =A.UNIT_CODE AND b.account_code = a.Vendor_code  and b.cust_drgno=a.cust_drgno AND B.trans_date = A.SCH_DATE AND B.Unit_Code='" + gstrUNITID + "' AND session_id='" & gstrIpaddressWinSck & "' and upper(ent_userid)='NAGARE COMP'  and a.kanbanno='TENTATIVE' )>0 AND A.UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.Open("select  B.cust_drgno,sum(B.quantity) as TotQty ,B.sch_date ,b.kanbanno from tmp_enagarodtl B Where B.Unit_Code='" + gstrUNITID + "' AND B.session_ID='" & gstrIpaddressWinSck & "' group by B.cust_drgno,B.sch_date ,b.kanbanno ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        CreateString = ""
        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                iQty = RsObjInsert.Fields("TotQty").Value 'scheduled quantity
                stracccode = Me.TxtCustCode.Text 'Account Code
                strSbuItCode = Trim(RsObjInsert.Fields("cust_drgno").Value) 'Cust_drgNo
                strKanbanNo = RsObjInsert.Fields("kanbanno").Value
                If RsObjItemcode.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjItemcode.Close()
                strItemCode = ""
                RsObjItemcode.Open("SELECT item_code FROM custitem_mst Where Unit_Code='" + gstrUNITID + "' AND cust_drgno='" & strSbuItCode & "' AND account_code='" & stracccode & "' and active=1", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjItemcode.EOF Then
                    strItemCode = Trim(RsObjItemcode.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                End If
                intYYYYMM = CInt(VB6.Format(RsObjInsert.Fields("sch_date").Value, "yyyymm")) 'Date in format YYYYMM
                'to check whether E nagare is already exists for the same date and for the same customer
                If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and status = 1  and Upper(ent_userid)='E NAGARE'") Then
                    strsql = " select * from tempitemcode Where Unit_Code='" + gstrUNITID + "' AND  tmp_accountcode ='" & stracccode & "' "
                    strsql = strsql & " and tmp_custdrgno ='" & strSbuItCode & "'   and tmp_itemcode ='" & strItemCode & "'"
                    strsql = strsql & " and tmp_transdate = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "'"
                    gobjDB.GetResult(strsql)
                    If gobjDB.RowCount >= 1 Then
                    Else
                        CreateString = CreateString & "'" & strSbuItCode & "'(" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "')  , "
                        strsql = " insert into tempitemcode(tmp_accountcode,tmp_custdrgno,tmp_itemcode,tmp_transdate,UNIT_CODE)values("
                        strsql = strsql & "'" & stracccode & "','" & strSbuItCode & "','" & strItemCode & "','"
                        strsql = strsql & RsObjInsert.Fields("sch_date").Value & "' ,'" + gstrUNITID + "')"
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                Else
                    'If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & Format(RsObjInsert.Fields("sch_date"), "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and status = 1 and (Upper(ent_userid) in('" & strKanbanNo & "'))") Then
                    If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and status = 1 ") Then
                        If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and status = 1 and (Upper(ent_userid) in('TENTATIVE'))") Then
                            strsql = "UPDATE DailyMKTSchedule_history Set RevisionNo=RevisionNo +1  Where Unit_Code='" + gstrUNITID + "' AND "
                            strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                            strsql = strsql & " and UPPER(ent_userid)='TENTATIVE'"
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            strsql = "Insert Into DailyMKTSchedule_history "
                            strsql = strsql & " select account_code,trans_date,item_code,cust_drgno,schedule_quantity,"
                            strsql = strsql & "  0, getdate() ,Ent_UserId,getdate(),Upd_UserId,UNIT_CODE from dailymktschedule where "
                            strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                            strsql = strsql & " AND UPPER(ent_userid)='TENTATIVE' AND Unit_Code='" + gstrUNITID + "'"
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            strsql = "delete from DailyMKTSchedule Where Unit_Code='" + gstrUNITID + "' AND "
                            strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                            strsql = strsql & " and ent_userid in('TENTATIVE') "
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                            strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE ) Values ( '" & stracccode & "', "
                            strsql = strsql & "'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, " dd mmm yyyy") & "', '"
                            strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                            strsql = strsql & CDbl(iQty) & " ,0 ,1"
                            strsql = strsql & ",0,getdate(),'" & RsObjInsert.Fields("Kanbanno").Value & "',getdate(),'" & RsObjInsert.Fields("Kanbanno").Value & "','" + gstrUNITID + "')"
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        Else
                            'IT MEANS THE ENTRY EXISTS FOR NAGARE COMP
                            If (RsObjInsert.Fields("Kanbanno")).Value = "NAGARE COMP" Then
                                strsql = "UPDATE DailyMKTSchedule_history Set RevisionNo=RevisionNo +1  Where Unit_Code='" + gstrUNITID + "' AND "
                                strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                                strsql = strsql & " and UPPER(ent_userid)='NAGARE COMP'"
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                strsql = "Insert Into DailyMKTSchedule_history "
                                strsql = strsql & " select account_code,trans_date,item_code,cust_drgno,schedule_quantity,"
                                strsql = strsql & "  0, getdate() ,Ent_UserId,getdate(),Upd_UserId,UNIT_CODE from dailymktschedule where "
                                strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                                strsql = strsql & " AND UPPER(ent_userid)='NAGARE COMP' AND Unit_Code='" + gstrUNITID + "'"
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                strsql = "delete from DailyMKTSchedule Where Unit_Code='" + gstrUNITID + "' AND "
                                strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "' AND Cust_drgno='" & strSbuItCode & "' "
                                strsql = strsql & " and ent_userid in('NAGARE COMP') "
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                                strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE ) Values ( '" & stracccode & "', "
                                strsql = strsql & "'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, " dd mmm yyyy") & "', '"
                                strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                                strsql = strsql & CDbl(iQty) & " ,0 ,1"
                                strsql = strsql & ",0,getdate(),'" & RsObjInsert.Fields("Kanbanno").Value & "',getdate(),'" & RsObjInsert.Fields("Kanbanno").Value & "','" + gstrUNITID + "')"
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                    Else
                        strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                        strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE ) Values ( '" & stracccode & "', "
                        strsql = strsql & "'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, " dd mmm yyyy") & "', '"
                        strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                        strsql = strsql & CDbl(iQty) & " ,0 ,1"
                        strsql = strsql & ",0,getdate(),'" & RsObjInsert.Fields("Kanbanno").Value & "',getdate(),'" & RsObjInsert.Fields("Kanbanno").Value & "','" + gstrUNITID + "')"
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
                RsObjInsert.MoveNext()
            End While
        End If
        If Len(CreateString) > 0 Then
            CreateString = Replace(CreateString, "'", "")
            CreateString = VB.Left(Trim(CreateString), Len(Trim(CreateString)) - 1)
            If MsgBox(" E-nagare is already uploaded for following items ( Date ) " & CreateString & vbCrLf & vbCrLf & " system will ignore to upload the above items " & vbCrLf & vbCrLf & " Do you want to continue to upload for remaining items ", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
            Else
                mP_Connection.RollbackTrans()
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
        End If
        'prashant rajpal changed ended for RFC requirement
        mP_Connection.CommitTrans()
        MsgBox("File has been uploaded successfully !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        txtDBFFilePath.Text = ""
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        If Err.Number = -2147217900 Then
            mP_Connection.RollbackTrans()
            MsgBox("Tentative Schedule Already Uploaded." & vbCrLf & Err.Description, MsgBoxStyle.Information, ResolveResString(100))
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        If Err.Number = -2147217913 Then
            mP_Connection.RollbackTrans()
            MsgBox("Invalid Selection." & vbCrLf & "Please Select Correct Schedule Option.", MsgBoxStyle.Information, ResolveResString(100))
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DISpares() As Object
        '-------------------------------------------------------------------------------------------------------------------------------------------
        ' Revision Date             -   27/03/2006
        ' Revision By               -   Davinder Singh
        ' Issue ID                  -   17378
        ' Revision History          -   1) DailyMKTSchedule table was not properly updated
        '                               2) Problem of linking of two Item_codes with same Cust_drgno of same customer solved
        '                               3) In SO Checking PO_Type('S') was not considered in some cases
        '                               4) Alow the same file uploaded repeatedly or cases where same KanBanNo repeates handled
        '                               5) Added the functionality by sending the data also in the Forecast_Mst
        '------------------------------------------------------------------------------------------------------------------------------------------
        ' Revision Date             -   02/06/2006
        ' Revision By               -   Davinder Singh
        ' Issue ID                  -   17995
        ' Revision History          -   To also check the newly added Po_Type='M' during SO checking
        '------------------------------------------------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim FSODISpares As Scripting.FileSystemObject
        Dim FSODISparesReadStatus As Scripting.TextStream
        Dim strstatus As String
        Dim i As Short
        Dim dblqty As Double
        Dim strMasterString As String
        Dim ArrMasterArray() As String
        Dim ArrSplitData() As String
        Dim stracccode As String
        Dim strItemCode As String
        Dim strSbuItCode As String
        Dim intYYYYMM As Integer
        Dim strsql As String
        Dim strSQLA As String
        Dim dblDispatchqty As Double
        Dim dblPrevSchedQty As Double
        Dim RsObjInsert As ADODB.Recordset
        Dim RsObjQuery As ADODB.Recordset
        Dim Rs As ADODB.Recordset
        Dim intMaxRNo As Short
        FSODISpares = New Scripting.FileSystemObject
        RsObjInsert = New ADODB.Recordset
        RsObjQuery = New ADODB.Recordset
        Rs = New ADODB.Recordset
        On Error GoTo ErrHandler
        mP_Connection.BeginTrans()
        FSODISparesReadStatus = FSODISpares.OpenTextFile(txtDBFFilePath.Text, Scripting.IOMode.ForReading, False)
        ''----Delete all data from temporary table Tmp_Enagarodtl for user's IP
        mP_Connection.Execute("DELETE FROM Tmp_Enagarodtl Where Unit_Code='" + gstrUNITID + "' AND Session_id='" & Trim(gstrIpaddressWinSck) & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''----Read the data from text file and put it into the temporary table Tmp_Enagarodtl
        Dim a As String
        While Not FSODISparesReadStatus.AtEndOfLine
            strMasterString = ""
            strstatus = FSODISparesReadStatus.ReadLine()
            ArrSplitData = Split(strstatus, " ")
            For i = 0 To UBound(ArrSplitData)
                If Len(Trim(ArrSplitData(i))) > 0 Then
                    strMasterString = strMasterString & ArrSplitData(i) & "»"
                End If
            Next
            ArrSplitData = Split(strMasterString, "»")
            If UBound(ArrSplitData) >= 8 Then
                strMasterString = ""
                For i = 0 To UBound(ArrSplitData)
                    'Changed for Issue ID eMpro-20090223-27780 Starts
                    'If i < 3 Or i > UBound(ArrSplitData) - 6 Then
                    If i <= 4 Or i > UBound(ArrSplitData) - 10 Then
                        'Changed for Issue ID eMpro-20090223-27780 Ends
                        If Len(Trim(ArrSplitData(i))) > 0 Then
                            strMasterString = strMasterString & ArrSplitData(i) & "»"
                        End If
                    End If
                Next
            End If
            'NAGARE and COMP are considered as two seperate words .So they are concatenated and inserted into the table
            ArrMasterArray = Split(strMasterString, "»")
            If UBound(ArrMasterArray) > 7 Then
                If IsDate(ArrMasterArray(1)) Then
                    If Len(ArrMasterArray(2)) = 5 Then
                        mP_Connection.RollbackTrans()
                        MsgBox("Invalid Schedule option Selected. File is E Nagare Schedule.", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Function
                    Else
                        'Commented for Issue ID eMpro-20090223-27780 Starts
                        'a = "INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Freq) values('" & Trim(Winsock.LocalIP) & "','" & Trim(ArrMasterArray(0)) & "','" & Trim(ArrMasterArray(2)) & "' ,'" & Trim(ArrMasterArray(3)) & "','" & Trim(ArrMasterArray(4)) & "',' ','" & Trim(ArrMasterArray(8)) & "','" & Trim(ArrMasterArray(1)) & "','23:59','" & ArrMasterArray(5) & "')"
                        'Commented for Issue ID eMpro-20090223-27780 Ends
                        mP_Connection.Execute("INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,Freq,UNIT_CODE,price_change_flag,batch_code,price) values('" & gstrIpaddressWinSck & "','" & Trim(ArrMasterArray(3)) & "','" & Trim(ArrMasterArray(4)) & "' ,'" & Trim(ArrMasterArray(5)) & "','" & Trim(ArrMasterArray(6)) & "','" & Trim(ArrMasterArray(8)) & "','" & Trim(ArrMasterArray(9)) & "','" & Trim(ArrMasterArray(1)) & "','11:59','" & ArrMasterArray(6) & "','" + gstrUNITID + "','" & Trim(ArrMasterArray(11)) & "' ,'" & Trim(ArrMasterArray(12)) & "','" & Trim(ArrMasterArray(13)) & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
            End If
        End While

        FSODISpares = Nothing
        FSODISparesReadStatus = Nothing
        ''----Fetch the whole data from temporary table for current IP to in a Recordset
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RsObjInsert.Open("SELECT * FROM Tmp_enagarodtl Where Unit_Code='" + gstrUNITID + "' AND Session_ID='" & Trim(gstrIpaddressWinSck) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                If Rs.State = ADODB.ObjectStateEnum.adStateOpen Then Rs.Close()
                'To retrieve Customer code line by line
                stracccode = ""
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                RsObjQuery.Open("SELECT customer_code FROM customer_mst Where Unit_Code='" + gstrUNITID + "' AND cust_vendor_code='" & Trim(RsObjInsert.Fields("vendor_code").Value) & "' and customer_code = '" & Trim(Me.TxtCustCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjQuery.EOF Then
                    stracccode = Trim(RsObjQuery.Fields("customer_code").Value)
                Else
                    MsgBox("No Data found in the Customer Master for the combination of seleted Customer Code[" & Trim(TxtCustCode.Text) & "] and customer vendor code[" & Trim(RsObjInsert.Fields("vendor_code").Value) & "] available in the file.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                ''----Pick item code from custitem_mst
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                RsObjQuery.Open("SELECT item_code FROM custitem_mst Where Unit_Code='" + gstrUNITID + "' AND  cust_drgno='" & RsObjInsert.Fields("cust_drgno").Value & "' AND account_code='" & stracccode & "' and active=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjQuery.EOF Then
                    strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                End If
                If RsObjQuery.RecordCount > 1 Then
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("SELECT D.Item_code from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE=D.UNIT_CODE AND H.Unit_Code='" + gstrUNITID + "' AND  H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type = 'M' and D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If Not RsObjQuery.EOF Then
                        strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                        GoTo Onerec
                    Else
                        If MsgBox(" There are more than 1 item code defined for this Customer part Code : " & Trim(RsObjInsert.Fields("cust_drgno").Value) & "." & vbCrLf & " Proceed with it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                            GoTo Onerec
                        Else
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        End If
                    End If
                ElseIf RsObjQuery.RecordCount < 1 Then  'Message for Item code is not Active and roll back the uploading
                    MsgBox(" Item Code not found for Customer Part Code code : " & Trim(RsObjInsert.Fields("cust_drgNo").Value) & vbCrLf & " Please correct the data first. It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                Else
Onerec:
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("select eNagareUploadingOnBasisOfSO FROM sales_parameter Where Unit_Code='" + gstrUNITID + "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If RsObjQuery.Fields(0).Value = True Then 'Value is set for eNagareUploadingOnBasisOfSO in sales_parameter
                        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                        RsObjQuery.Open("SELECT D.Item_code from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE=D.UNIT_CODE AND H.Unit_Code='" + gstrUNITID + "' AND H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 AND D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If RsObjQuery.RecordCount = 0 Then 'There is no SO Active & Authorized
                            MsgBox(" There is no SO Authorized or Active for Cust Part Code: " & Trim(RsObjInsert.Fields("cust_drgNo").Value) & " for selected Customer. " & vbCrLf & " It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        Else
                            strItemCode = Trim(RsObjQuery.Fields("Item_Code").Value)
                        End If
                    End If
                End If
                '----If current kanban already exist then read its Qty. and Reduce it from the respective quantities of DailyMKTSchedule, forecast_mst and delete from mkt_enagaredtl
                strSQLA = "select Quantity from mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' AND Account_code = '" & stracccode & "' and Item_code = '" & strItemCode & "' and Cust_drgno = '" & RsObjInsert.Fields("cust_drgno").Value & "' and kanbanno = '" & RsObjInsert.Fields("kanbanno").Value & "'"
                Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                Rs.Open(strSQLA, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Rs.RecordCount >= 1 Then
                    If ValidateData("Item_Code", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_DrgNo = '" & RsObjInsert.Fields("cust_drgno").Value & "' And Item_Code = '" & strItemCode & "'") Then
                        strsql = "UPDATE DailyMKTSchedule Set Schedule_quantity=Schedule_quantity-" & Val(Rs.Fields("Quantity").Value) & ", Upd_UserId = 'DI SPARES', Upd_dt = getdate() Where Unit_Code='" + gstrUNITID + "' AND  Status = 1 and"
                        strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_Code = '" & strItemCode & "'"
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If Len(Find_Value("select product_no from forecast_mst Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' and Enagare_Unloc='" & RsObjInsert.Fields("unloc").Value & "'")) > 0 Then
                        mP_Connection.Execute("Update forecast_mst Set Quantity=Quantity -" & Val(Rs.Fields("Quantity").Value) & ",Upd_dt=Getdate(),upd_userid='DI SPARES'  Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' and Enagare_UNLOC='" & RsObjInsert.Fields("unloc").Value & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    mP_Connection.Execute("DELETE FROM mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' AND kanbanno='" & Trim(RsObjInsert.Fields("kanbanno").Value) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                ''----Insert the record containing new KanbanNo into the table mkt_enagaredtl
                mP_Connection.Execute("Insert Into mkt_enagaredtl(Account_code,Item_code,Cust_drgno,Quantity,UNLOC,USLOC,KanbanNo,Sch_date,Sch_time,scheduletype,UNIT_CODE,price_change_flag,batch_code,price) VALUES ( '" & stracccode & "' ,'" & strItemCode & "','" & RsObjInsert.Fields("cust_drgno").Value & "','" & RsObjInsert.Fields("quantity").Value & "','" & RsObjInsert.Fields("unloc").Value & "','" & RsObjInsert.Fields("usloc").Value & "','" & RsObjInsert.Fields("kanbanno").Value & "','" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "','" & RsObjInsert.Fields("sch_time").Value & "','M','" + gstrUNITID + "','" & RsObjInsert("price_change_flag").Value & "','" & RsObjInsert("batch_code").Value & "','" & RsObjInsert("price").Value & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                ''----If entry for current record already exist in forecast_mst then update it else insert new entry
                If Len(Find_Value("select product_no from forecast_mst Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' and Enagare_Unloc='" & RsObjInsert.Fields("unloc").Value & "' AND ent_userid='DI SPARES'")) > 0 Then
                    mP_Connection.Execute("Update forecast_mst Set Quantity=Quantity +" & Val(RsObjInsert.Fields("quantity").Value) & ",Upd_dt=Getdate(),upd_userid='DI SPARES'  Where Unit_Code='" + gstrUNITID + "' AND customer_code='" & stracccode & "' and product_no='" & strItemCode & "' and due_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' and Enagare_UNLOC='" & RsObjInsert.Fields("unloc").Value & "' AND ent_userid='DI SPARES'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Else
                    strsql = "INSERT INTO forecast_mst(Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt, ENagare_UNLOC,UNIT_CODE) VALUES ('" & stracccode & "','" & strItemCode & "' ,'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "'," & Val(RsObjInsert.Fields("quantity").Value) & ",'DI SPARES',getdate(),'DI SPARES',getdate(), '" & RsObjInsert.Fields("unloc").Value & "','" + gstrUNITID + "')"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                RsObjInsert.MoveNext()
            End While
        Else
            mP_Connection.RollbackTrans()
            MsgBox("No data Found for insertion", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        End If
        ''----To read data from tables cust_ord_hdr, cust_ord_dtl, customer_mst, tmp_enagarodtl to insert into DailyMKTschedule table
        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
        RsObjQuery.Open("select  A.Account_code,A.item_code,B.cust_drgno ,sum(B.quantity) as TotQty ,B.sch_date from mkt_enagaredtl A, tmp_enagarodtl B WHERE A.UNIT_CODE=B.UNIT_CODE AND A.Unit_Code='" + gstrUNITID + "' AND  A.KanbanNo = B.KanbanNo and A.Cust_Drgno = B.Cust_Drgno and B.session_ID='" & gstrIpaddressWinSck & "' group by A.Account_code,A.item_code,B.cust_drgno,B.sch_date ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjQuery.EOF Then
            While Not RsObjQuery.EOF
                dblqty = Val(RsObjQuery.Fields("TotQty").Value)
                stracccode = Trim(TxtCustCode.Text)
                strSbuItCode = RsObjQuery.Fields("cust_drgNo").Value
                strItemCode = RsObjQuery.Fields("Item_Code").Value
                intYYYYMM = CInt(VB6.Format(RsObjQuery.Fields("sch_date").Value, "yyyymm")) 'Date in format YYYYMM
                ''----To check if Record already exist in the DailyMKTSchedule or not
                If ValidateData("Item_Code", "DailyMKTSchedule", "Account_code='" & stracccode & "' AND Trans_date='" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' ") Then
                    ''----Item exist in DailyMKTSchedule so delete from MonthlyMKTSchedule
                    strsql = " Delete From MonthlyMKTSchedule Where Unit_Code='" + gstrUNITID + "' AND Account_Code = '" & stracccode & "'"
                    strsql = strsql & " And Cust_DrgNo = '" & strSbuItCode & "' AND Item_Code = '" & strItemCode & "'"
                    strsql = strsql & " AND Status = 1 AND Year_Month = " & intYYYYMM
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    ''----Read despatch from DailyMKTSchedule
                    dblDispatchqty = Val(SelectDataFromTable("Despatch_Qty", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status=1"))
                    ''----Read Schedule from DailyMKTSchedule
                    dblPrevSchedQty = Val(SelectDataFromTable("Schedule_Quantity", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status=1"))
                    ''----Read Max Revision No. from DailyMKTSchedule
                    '10565410  
                    'intMaxRNo = CShort(SelectDataFromTable("RevisionNo", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1"))
                    intMaxRNo = CShort(Val(SelectDataFromTable("RevisionNo", "DailyMKTSchedule", " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1")))
                    '10565410  end 
                    ''----Update DailyMKTSchedule by incrementing revision No. by 1 and setting status=0
                    strsql = "UPDATE DailyMKTSchedule set Status = 0, Upd_UserId = 'DI SPARES', Upd_dt = getdate() Where Unit_Code='" + gstrUNITID + "' AND "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_COde = '" & strItemCode & "' and status = 1"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    ''----Insert new record with Revision No. = Max(RevisionNo)+1 and status=1
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE ) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & " ', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    strsql = strsql & dblqty + dblPrevSchedQty & " ," & dblDispatchqty & " ,1"
                    strsql = strsql & "," & intMaxRNo + 1 & ",getdate(),'DI SPARES',getdate(),'DI SPARES','" + gstrUNITID + "' )"
                Else ''----Entry does't exist in the DailyMKTSchedule
                    ''----Insert new record with Revision No.= 0 and status = 1
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId ,UNIT_CODE) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjQuery.Fields("sch_date").Value, "dd mmm yyyy") & "', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    strsql = strsql & dblqty & " ,0 ,1"
                    strsql = strsql & ",0,getdate(),'DI SPARES',getdate(),'DI SPARES','" + gstrUNITID + "' )"
                End If
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                RsObjQuery.MoveNext()
            End While
        End If
        mP_Connection.CommitTrans()
        MsgBox("File has been uploaded successfully !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        txtDBFFilePath.Text = ""
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
        If Rs.State = ADODB.ObjectStateEnum.adStateOpen Then Rs.Close()
        RsObjInsert = Nothing
        RsObjQuery = Nothing
        Rs = Nothing
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        If Err.Number = -2147217900 Then
            mP_Connection.RollbackTrans()
            MsgBox("Data already uploaded. Quitting the job", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        End If
        If Err.Number = -2147217833 Then
            mP_Connection.RollbackTrans()
            MsgBox("Invalid Schedule Selection. Quitting the job", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    'added for sun vacuum since kanban no can be repeated in the next file
    Private Function ChkNagarroNo(ByVal strnagno As String) As Boolean
        Dim RsObjNagarroNo As New ADODB.Recordset
        On Error GoTo ErrHandler
        ChkNagarroNo = False
        If RsObjNagarroNo.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjNagarroNo.Close()
        RsObjNagarroNo.Open("SELECT kanbanno FROM mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' AND kanbanno='" & strnagno & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjNagarroNo.EOF Then
            ChkNagarroNo = True
        End If
        If RsObjNagarroNo.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjNagarroNo.Close()
        RsObjNagarroNo = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
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
    Private Sub optDIQuery_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDIQuery.CheckedChanged
        '----------------------------------------------------------------------------
        'Author         :   Manoj Vaish
        'Function       :   Visible the option to select SO Type
        'Comments       :   Issue ID eMpro-20090505-31005
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If optDIQuery.Checked = True Then
            lblSOtype.Visible = True
            cmbsotype.Visible = True
            cmbsotype.SelectedIndex = -1
        Else
            lblSOtype.Visible = False
            cmbsotype.Visible = False
            cmbsotype.SelectedIndex = -1
        End If
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class