'-------------------------------------------------------------
'Created By     : Shabbir Hussain
'Created On     : 11 July 2011
'Description    : TKML PDS and Receipt Uploading  
'Issue ID       : TT# 10113522
'-------------------------------------------------------------
' Revised By                 -   Roshan Singh
' Revision Date              -   11 Oct 2011
' Description                -   FOR MULTIUNIT FUNCTIONALITY and CHANGE MANAGEMENT
'-------------------------------------------------------------
'Revised by     : Prashant Rajpal
'Issue Id       : 10162423 
'Description    : Unable to upload the PDS through eMPro 
'-------------------------------------------------------------
'Modified By Roshan Singh on 19 Dec 2011 for multiUnit change management    
'-------------------------------------------------------------
'Revised by     : Abhinav Kumar
'Issue Id       : 10715594  
'Description    : New functionality for Auto-Mailer in PDS Material Receipt
'---------------------------------------------------------------------------------
'Created By     : Parveen Kumar
'Created On     : 13 FEB 2015
'Description    : eMPro Vehicle BOM
'Issue ID       : 10737738 
'----------------------------------------------------------------------------------

Imports System.Data
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class FRMMKTTRN0072
#Region "Form level variable Declarations"
    Dim mintFormtag As Short
    Dim mstrIPAddress As String
    Dim mctlError As System.Windows.Forms.Control
    Dim SchUpdFlag As Boolean = False   '10737738
    Dim SchPDSCUSTUpdFlag As Boolean = False   '10737738

    Private Enum enumsspr
        Err_desc = 1
        Source
    End Enum
#End Region
#Region "Routines"
    Private Sub SetSpreadProperty()
        Try
            With sspr
                .DisplayRowHeaders = True
                .MaxRows = 0
                .MaxCols = 0
                .MaxCols = enumsspr.Source
                .UserColAction = FPSpreadADO.UserColActionConstants.UserColActionSort
                .Row = 0
                .Col = enumsspr.Err_desc : .Text = "Error Description" : .set_ColWidth(enumsspr.Err_desc, 3500)
                .Col = enumsspr.Source : .Text = "Source" : .set_ColWidth(enumsspr.Source, 8100)
                .set_RowHeight(.Row, 400)
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    'Changes against 10737738 
    Private Sub ChkVBSchUpdFlag()
        Dim strSql As String = String.Empty

        Try

            strSql = " select top 1 1 from sales_parameter where Unit_Code='" & gstrUnitId & "' and SCHEDULE_UPLOAD_CONFIG = 1  "
            SchUpdFlag = IsRecordExists(strSql)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub ChkPDSCUST_UpdFlag()
        Dim strSql As String = String.Empty
        Try
            strSql = " select top 1 1 from CUSTOMER_MST where Unit_Code='" & gstrUnitId & "' and PDS_UPLOAD_SECONDARYPARTCODE_FLAG = 1 and customer_code='" & TxtCustomerCode.Text.Trim & "'"
            SchPDSCUSTUpdFlag = IsRecordExists(strSql)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub SetSpreadColTypes(ByVal pintRowNo As Integer)
        Try
            With sspr
                .Row = pintRowNo
                .Col = enumsspr.Err_desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Source : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub addRowAtEnterKeyPress(ByVal pintRows As Integer)
        Dim intRowHeight As Integer
        Try
            With sspr
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                For intRowHeight = 1 To pintRows
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .set_RowHeight(.Row, 350)
                    Call SetSpreadColTypes(.Row)
                Next intRowHeight
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Function PopulateScheduleLog() As Boolean
        Dim objconn As SqlConnection = Nothing
        Dim objDR As SqlDataReader
        Dim objCommand As New SqlCommand()
        PopulateScheduleLog = False
        SetSpreadProperty()
        Try
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = "SELECT ERR_DESC,SOURCE FROM SCHEDULE_UPLOAD_LOG WITH (NOLOCK) WHERE UNIT_CODE = '" & gstrUnitId & "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' ORDER BY LOG_ID "
            objCommand.CommandType = CommandType.Text
            objDR = objCommand.ExecuteReader()
            If objDR.HasRows = True Then
                While objDR.Read
                    addRowAtEnterKeyPress(1)
                    With sspr
                        .Row = .MaxRows
                        .Col = enumsspr.Err_desc
                        .Text = objDR("ERR_DESC").ToString.Trim
                        .TypeTextWordWrap = True
                        .Col = enumsspr.Source
                        .Text = objDR("SOURCE").ToString.Trim
                        .TypeTextWordWrap = True
                    End With
                End While
                sspr.Visible = True
                PopulateScheduleLog = True
            End If
            If objDR.IsClosed = False Then objDR.Close()
            objDR = Nothing
            objCommand = Nothing
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objconn = Nothing
        Catch ex As Exception
            objDR = Nothing
            objCommand = Nothing
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objconn = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function

    Private Function ReadMarutiPDSFileProc() As Boolean
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
        Dim strcustvendcode As String 'Customer vendor Code
        Dim strcustdrgno As String 'Customer Drg No
        Dim strPDSNo As String 'PDS No
        Dim strSchedDate As Date
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
        Dim objCom As ADODB.Command
        Rs = New ADODB.Recordset
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        On Error Resume Next
        mP_Connection.Execute("DELETE FROM tempdata WHERE Unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        On Error GoTo ErrHandler
        mP_Connection.BeginTrans()

        ''----Read the data from textfile and put it into the temporary table Tmp_Enagarodtl
        FSOReadStatus = FSODIDchedules.OpenTextFile(lblFileName.Text, Scripting.IOMode.ForReading, False)
        mP_Connection.Execute("DELETE FROM Tmp_Enagarodtl WHERE Session_id='" & gstrIpaddressWinSck & "' and UNIT_CODE = '" & gstrUNITID & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

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
                    Dim strDate As Date = ArrMasterArray(2)
                    Dim strDelDate As Date = ArrMasterArray(1)
                    If strDate < GetServerDate() Then
                        MsgBox("Collection date is less than current date for drawing No " & ArrMasterArray(4) & ". " & vbCrLf & " It should be equal or greater than current date . " & vbCrLf & " File cannot be upload  for selected Customer. ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        mP_Connection.RollbackTrans()
                        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        Exit Function
                    ElseIf strDelDate < GetServerDate() Then
                        MsgBox("Delivery date is less than current date for drawing No " & ArrMasterArray(4) & ". " & vbCrLf & " It should be equal or greater than current date . " & vbCrLf & " File cannot be upload  for selected Customer. ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        mP_Connection.RollbackTrans()
                        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        Exit Function
                    End If
                    'SESSION ID » VENDOR CODE » DEL DATE » SCH DATE  » SCH TIME »  CUSTDRGNO » QTY » PDS NO      ¬
                    Dim str As String = "INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,DelDate,Sch_date,Sch_time,Cust_drgno,Quantity,KanbanNo,UNIT_CODE,UNLOC,USLOC) values('" & gstrIpaddressWinSck & "','" & Trim(ArrMasterArray(0)) & "','" & Trim(ArrMasterArray(1)) & "','" & Trim(ArrMasterArray(2)) & "' ,'" & Trim(ArrMasterArray(3)) & "','" & Trim(ArrMasterArray(4)) & "','" & Trim(ArrMasterArray(5)) & "','" & Trim(ArrMasterArray(6)) & "' + ' ' + '" & Trim(ArrMasterArray(7)) & "','" & gstrUNITID & "','','')"
                    mP_Connection.Execute(str, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
        End While



        'mP_Connection.Execute("Delete from Tmp_enagarodtl where UNIT_CODE = '" & gstrUNITID & "' and kanbanNo like 'Nagare%'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''----To read the data from Tmp_enagarodtl and put it into the Recordset
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RsObjInsert.Open("SELECT * FROM Tmp_enagarodtl where UNIT_CODE = '" & gstrUNITID & "' and Session_ID='" & gstrIpaddressWinSck & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'prasha nt 
        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                iQty = RsObjInsert.Fields("Quantity").Value 'Cumulative quantity to be inserted to daily marketting schedule of that item
                stracccode = Trim(TxtCustomerCode.Text) 'Account Code
                strcustdrgno = Trim(RsObjInsert.Fields("cust_drgno").Value) 'Cust_drgNo
                strcustvendcode = Trim(RsObjInsert.Fields("Vendor_code").Value) 'Cust_drgNo
                strPDSNo = Trim(RsObjInsert.Fields("KanbanNo").Value) 'Cust_drgNo
                strSchedDate = VB6.Format(RsObjInsert.Fields("Sch_date").Value, "dd mmm yyyy") 'Cust_drgNo
                intYYYYMM = CInt(VB6.Format(RsObjInsert.Fields("sch_date").Value, "yyyymm"))
                objCom = New ADODB.Command
                With objCom
                    .ActiveConnection = mP_Connection
                    .CommandTimeout = 0
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "USP_UPLOAD_PDSMARUTI_SCHEDULE"
                    .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    .Parameters.Append(.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 12, Trim(TxtCustomerCode.Text)))
                    .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, gstrIpaddressWinSck.Trim()))
                    .Parameters.Append(.CreateParameter("@USERID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, mP_User.Trim()))
                    .Parameters.Append(.CreateParameter("@CUST_VENDOR_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 12, Trim(strcustvendcode)))
                    .Parameters.Append(.CreateParameter("@CUST_DRGNO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(strcustdrgno)))
                    .Parameters.Append(.CreateParameter("@PDSNo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, Trim(strPDSNo)))
                    .Parameters.Append(.CreateParameter("@DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, strSchedDate))
                    .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 1000))


                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                End With
                If objCom.Parameters(objCom.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                    mP_Connection.RollbackTrans()
                    MessageBox.Show(objCom.Parameters(objCom.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    objCom = Nothing
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                RsObjInsert.MoveNext()
            End While
        End If

        mP_Connection.CommitTrans()
        MsgBox("File has been uploaded successfully !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        lblFileName.Text = ""
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        If Err.Number = -2147217833 Then
            mP_Connection.RollbackTrans()
            MsgBox("Invalid Schedule Selection." & vbCrLf & "Please Select Correct Schedule Option.", MsgBoxStyle.Information, ResolveResString(100))
            lblFileName.ForeColor = System.Drawing.Color.Red
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        If Err.Number = -2147217900 Then
            mP_Connection.RollbackTrans()
            MsgBox("Data already uploaded", MsgBoxStyle.Information, ResolveResString(100))
            lblFileName.ForeColor = System.Drawing.Color.Red
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function ReadMarutiPDSFile() As Boolean
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
        On Error Resume Next
        mP_Connection.Execute("DELETE FROM tempdata WHERE Unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        On Error GoTo ErrHandler
        mP_Connection.BeginTrans()

        ''----Read the data from textfile and put it into the temporary table Tmp_Enagarodtl
        FSOReadStatus = FSODIDchedules.OpenTextFile(lblFileName.Text, Scripting.IOMode.ForReading, False)
        mP_Connection.Execute("DELETE FROM Tmp_Enagarodtl WHERE Session_id='" & gstrIpaddressWinSck & "' and UNIT_CODE = '" & gstrUNITID & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

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
                    Dim strDate As Date = ArrMasterArray(2)
                    Dim strDelDate As Date = ArrMasterArray(1)
                    If strDate < GetServerDate() Then
                        MsgBox("Collection date is less than current date for drawing No " & ArrMasterArray(4) & ". " & vbCrLf & " It should be equal or greater than current date . " & vbCrLf & " File cannot be upload  for selected Customer. ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        mP_Connection.RollbackTrans()
                        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        Exit Function
                    ElseIf strDelDate < GetServerDate() Then
                        MsgBox("Delivery date is less than current date for drawing No " & ArrMasterArray(4) & ". " & vbCrLf & " It should be equal or greater than current date . " & vbCrLf & " File cannot be upload  for selected Customer. ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        mP_Connection.RollbackTrans()
                        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        Exit Function
                    End If
                    'SESSION ID » VENDOR CODE » DEL DATE » SCH DATE  » SCH TIME »  CUSTDRGNO » QTY » PDS NO      ¬
                    Dim str As String = "INSERT INTO Tmp_Enagarodtl(Session_ID,vendor_code,DelDate,Sch_date,Sch_time,Cust_drgno,Quantity,KanbanNo,UNIT_CODE,UNLOC,USLOC) values('" & gstrIpaddressWinSck & "','" & Trim(ArrMasterArray(0)) & "','" & Trim(ArrMasterArray(1)) & "','" & Trim(ArrMasterArray(2)) & "' ,'" & Trim(ArrMasterArray(3)) & "','" & Trim(ArrMasterArray(4)) & "','" & Trim(ArrMasterArray(5)) & "','" & Trim(ArrMasterArray(6)) & "' + ' ' + '" & Trim(ArrMasterArray(7)) & "','" & gstrUNITID & "','','')"
                    mP_Connection.Execute(str, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
        End While

   
        'mP_Connection.Execute("Delete from Tmp_enagarodtl where UNIT_CODE = '" & gstrUNITID & "' and kanbanNo like 'Nagare%'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''----To read the data from Tmp_enagarodtl and put it into the Recordset
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RsObjInsert.Open("SELECT * FROM Tmp_enagarodtl where UNIT_CODE = '" & gstrUNITID & "' and Session_ID='" & gstrIpaddressWinSck & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'prashant 
        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
        RsObjQuery.Open("select isnull(ENAGARE_ALLOWED_ALREADYINVOICE,0)as ENAGARE_ALLOWED_ALREADYINVOICE  FROM sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "' ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If RsObjQuery.Fields(0).Value = True Then 'Value is set for eNagareUploadingOnBasisOfSO in sales_parameter
            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
            RsObjQuery.Open("SELECT KANBANNO FROM VW_ENAGAREUPLOAD_ALREADYINVOICEPDS_MARUTI where UNIT_CODE = '" & gstrUNITID & "' and Session_ID='" & gstrIpaddressWinSck & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If RsObjQuery.RecordCount > 0 Then 'There is no SO Active & Authorized
                For intloopcounter = 1 To RsObjQuery.RecordCount
                    If strkanbanno = "" Then
                        strkanbanno = RsObjQuery.Fields("KANBANNO").Value.ToString
                    Else
                        strkanbanno = strkanbanno & "," & RsObjQuery.Fields("KANBANNO").Value.ToString
                    End If
                Next
                MsgBox("The Invoice for PDS Nos. " & strkanbanno & "  has already been generated. File cannot be upload  for selected Customer. " & vbCrLf & " It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                mP_Connection.RollbackTrans()
                ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
        End If

        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                'To retrieve Customer code line by line
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                stracccode = ""
                RsObjQuery.Open("SELECT customer_code FROM customer_mst WHERE UNIT_CODE = '" & gstrUNITID & "' and cust_vendor_code='" & Trim(RsObjInsert.Fields("vendor_code").Value) & "' and customer_code = '" & Trim(Me.TxtCustomerCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjQuery.EOF Then
                    stracccode = Trim(RsObjQuery.Fields("customer_code").Value)
                Else
                    MsgBox("No Data found in the Customer Master for the combination of seleted Customer Code[" & Trim(TxtCustomerCode.Text) & "] and customer vendor code[" & Trim(RsObjInsert.Fields("vendor_code").Value) & "] in the file.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                ''----Read the item_code from the custitem_mst table
                If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                strItemCode = ""
                RsObjQuery.Open("SELECT item_code FROM custitem_mst WHERE UNIT_CODE = '" & gstrUNITID & "' and cust_drgno='" & Trim(RsObjInsert.Fields("Cust_drgno").Value) & "' AND account_code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not RsObjQuery.EOF Then
                    strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                End If
                'Changed for More than one item code active for more
                'than one SO authorized and active then pick depending on Sales_parameter
                If RsObjQuery.RecordCount < 1 Then  'Message for Item code is not Active and roll back the uploading
                    MsgBox(" Item Code not found for Cust Part Code code : " & Trim(RsObjInsert.Fields("Cust_drgno").Value) & vbCrLf & " Please correct the data first. It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    mP_Connection.RollbackTrans()
                    ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Exit Function
                ElseIf RsObjQuery.RecordCount > 1 Then
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("SELECT D.Item_code from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE = D.UNIT_CODE AND H.UNIT_CODE = '" & gstrUNITID & "' AND H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type = 'O' and D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If Not RsObjQuery.EOF Then
                        strItemCode = Trim(RsObjQuery.Fields("Item_code").Value) 'Item code is Fetched to be inserted into the table MKT_EnagareDtl as it was working previously
                        GoTo Onerec
                    Else
                        If MsgBox("There are more than 1 item code defined for this Customer part Code : " & Trim(RsObjInsert.Fields("cust_drgno").Value) & "." & vbCrLf & " Proceed with it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        Else
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        End If
                    End If
                Else 'There is only one row / Item Code for the defined Internal Code
                    'Code Added by Arshad on 05/04/2005
Onerec:
                    If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                    RsObjQuery.Open("select eNagareUploadingOnBasisOfSO FROM sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "' ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If RsObjQuery.Fields(0).Value = True Then 'Value is set for eNagareUploadingOnBasisOfSO in sales_parameter
                        If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                        'RsObjQuery.Open("SELECT D.Item_code from cust_ord_hdr H , cust_ord_Dtl D WHERE H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type='O' and D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        RsObjQuery.Open("SELECT  D.Item_code,H.PO_TYPE from cust_ord_hdr H , cust_ord_Dtl D WHERE H.UNIT_CODE = D.UNIT_CODE and H.UNIT_CODE = '" & gstrUNITID & "' AND  H.Account_Code=D.Account_Code and H.Cust_Ref=D.Cust_Ref   and H.Amendment_No=D.Amendment_No and H.Authorized_Flag=1 and H.po_type ='O' and D.Active_Flag='A' and D.cust_drgNo='" & Trim(RsObjInsert.Fields("cust_drgNo").Value) & "' AND D.Account_Code='" & stracccode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If RsObjQuery.RecordCount = 0 Then 'There is no SO Active & Authorized
                            MsgBox(" There is no SO Authorized and Active for Item " & Trim(RsObjInsert.Fields("Cust_Drgno").Value) & " for selected Customer. " & vbCrLf & " It will cancel the schedule uploading", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            If RsObjQuery.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjQuery.Close()
                            mP_Connection.RollbackTrans()
                            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Exit Function
                        Else 'Item code is Fetched to be inserted into the table MKT_EnagareDtl
                            strItemCode = Trim(RsObjQuery.Fields("Item_Code").Value)
                        End If
                    End If
                End If
                '----case specifically added for sun vacuum--if same nagarro nos is repeated,then
                'delete the previous nos from mktenagaro_dtl and insert new one .
                If ChkNagarroNo(RsObjInsert.Fields("kanbanno").Value, strItemCode) Then
                    blnchangeddata = True
                    If blnchangeddata = True Then
                        mP_Connection.Execute("DELETE FROM mkt_enagaredtl WHERE kanbanno='" & RsObjInsert.Fields("kanbanno").Value & "' and Item_code='" & strItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("Insert Into mkt_enagaredtl(Account_code,Item_code,Cust_drgno,Quantity,KanbanNo,Sch_date,Sch_time,Scheduletype,UNIT_CODE,DelDate,UNLOC,USLOC) VALUES ( '" & stracccode & "' ,'" & strItemCode & "','" & RsObjInsert.Fields("cust_drgno").Value & "','" & RsObjInsert.Fields("quantity").Value & "','" & RsObjInsert.Fields("kanbanno").Value & "','" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "','" & RsObjInsert.Fields("sch_time").Value & "','O','" & gstrUNITID & "','" & VB6.Format(RsObjInsert.Fields("DelDate").Value, "dd mmm yyyy") & "','','')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                Else  'nagare no not exists 
                    mP_Connection.Execute("Insert Into mkt_enagaredtl(Account_code,Item_code,Cust_drgno,Quantity,KanbanNo,Sch_date,Sch_time,Scheduletype,UNIT_CODE,DelDate,UNLOC,USLOC) VALUES ( '" & stracccode & "' ,'" & strItemCode & "','" & RsObjInsert.Fields("cust_drgno").Value & "','" & RsObjInsert.Fields("quantity").Value & "','" & RsObjInsert.Fields("kanbanno").Value & "','" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "','" & RsObjInsert.Fields("sch_time").Value & "','O','" & gstrUNITID & "','" & VB6.Format(RsObjInsert.Fields("Deldate").Value, "dd mmm yyyy") & "','','')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                End If 'nagare no loop closed
        RsObjInsert.MoveNext()
            End While
        Else
            mP_Connection.RollbackTrans()
            MsgBox("No data Found for insertion", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        If RsObjInsert.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjInsert.Close()
        RsObjInsert.Open("select  A.Account_code,A.item_code,B.cust_drgno ,sum(B.quantity) as TotQty ,B.sch_date from mkt_enagaredtl A, tmp_enagarodtl B where A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "' AND A.KanbanNo = B.KanbanNo and A.Cust_Drgno = B.Cust_Drgno and B.session_ID='" & gstrIpaddressWinSck & "' group by A.Account_code,A.item_code,B.cust_drgno,B.sch_date ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjInsert.EOF Then
            While Not RsObjInsert.EOF
                iQty = RsObjInsert.Fields("TotQty").Value 'Cumulative quantity to be inserted to daily marketting schedule of that item
                stracccode = Trim(RsObjInsert.Fields("Account_code").Value) 'Account Code
                strSbuItCode = Trim(RsObjInsert.Fields("cust_drgno").Value) 'Cust_drgNo
                strItemCode = Trim(RsObjInsert.Fields("Item_Code").Value) 'Item code
                intYYYYMM = CInt(VB6.Format(RsObjInsert.Fields("sch_date").Value, "yyyymm")) 'Date in format YYYYMM
             
                'prashant rajpal changed ended as per RFC
                If ValidateData("Item_Code", "DailyMKTSchedule", " Account_code='" & stracccode & "' and UNIT_CODE = '" & gstrUNITID & "' AND Trans_date='" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Cust_drgno='" & strSbuItCode & "' AND item_code='" & strItemCode & "' and status = 1  ") Then
                    'If Item HAS Monthly Schedule then delete it
                    strsql = " Delete From MonthlyMKTSchedule Where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code = '" & stracccode & "'"
                    strsql = strsql & " And Cust_DrgNo = '" & strSbuItCode & "' AND Item_Code = '" & strItemCode & "'"
                    strsql = strsql & " AND Status = 1 AND Year_Month = " & intYYYYMM
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    'If Entry for this Item is already there .... for strFields(3) Date
                    dblDispatchqty = CDbl(Val(SelectDataFromTable("Despatch_Qty", "DailyMKTSchedule", "  UNIT_CODE = '" & gstrUNITID & "' and Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1")))
                    dblPrevSchedQty = CDbl(Val(SelectDataFromTable("Schedule_Quantity", "DailyMKTSchedule", "  UNIT_CODE = '" & gstrUNITID & "' and Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1")))
                    intMaxRNo = CShort(SelectDataFromTable("RevisionNo", "DailyMKTSchedule", "   UNIT_CODE = '" & gstrUNITID & "' and Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "' AND Status = 1"))
                    '1.)Update and INCREASE by 1 Revision No,Status = 0 for Previous Entries
                    strsql = "UPDATE DailyMKTSchedule Set Status = 0, schedule_flag=0 ,Upd_Userid='PDS Nagare',upd_dt=getdate() Where "
                    strsql = strsql & " Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' AND Item_COde = '" & strItemCode & "' and status = 1 and UNIT_CODE = '" & gstrUNITID & "'"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId ,UNIT_CODE) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & " ', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    strsql = strsql & dblPrevSchedQty + iQty & " ," & dblDispatchqty & " ,1"
                    strsql = strsql & "," & intMaxRNo + 1 & ",getdate(),'PDS Nagare',getdate(),'PDS Nagare','" & gstrUNITID & "' )"
                Else
                    'Insert Item into DailyMKTSchedule
                    dblDispatchqty = CDbl(Val(SelectDataFromTable("Despatch_Qty", "DailyMKTSchedule", "  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code = '" & stracccode & "' AND Trans_Date = '" & VB6.Format(RsObjInsert.Fields("sch_date").Value, "dd mmm yyyy") & "' And Item_Code = '" & strItemCode & "'")))
                    strsql = "Insert Into DailyMKTSchedule (Account_Code,Trans_date,Item_code,Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,"
                    strsql = strsql & " Status,RevisionNo, Ent_dt,Ent_UserId,Upd_dt,Upd_UserId ,UNIT_CODE) Values ( '" & stracccode & "', "
                    strsql = strsql & "'" & VB6.Format(RsObjInsert.Fields("sch_date").Value, " dd mmm yyyy") & "', '"
                    strsql = strsql & strItemCode & "', '" & strSbuItCode & "',1, "
                    'strsql = strsql & CDbl(iQty) & " ,0 ,1"
                    strsql = strsql & CDbl(iQty) & " ," & dblDispatchqty & ",1"
                    strsql = strsql & ",0,getdate(),'PDS Nagare',getdate(),'PDS Nagare','" & gstrUNITID & "' )"
                End If
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                RsObjInsert.MoveNext()
            End While
        End If
        mP_Connection.CommitTrans()
        MsgBox("File has been uploaded successfully !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        lblFileName.Text = ""
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        If Err.Number = -2147217833 Then
            mP_Connection.RollbackTrans()
            MsgBox("Invalid Schedule Selection." & vbCrLf & "Please Select Correct Schedule Option.", MsgBoxStyle.Information, ResolveResString(100))
            lblFileName.ForeColor = System.Drawing.Color.Red
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        If Err.Number = -2147217900 Then
            mP_Connection.RollbackTrans()
            MsgBox("Data already uploaded", MsgBoxStyle.Information, ResolveResString(100))
            lblFileName.ForeColor = System.Drawing.Color.Red
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Function
        End If
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
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
    Private Function ChkNagarroNo(ByVal strnagno As String, ByVal strItem As String) As Boolean
        Dim RsObjNagarroNo As New ADODB.Recordset
        On Error GoTo ErrHandler
        ChkNagarroNo = False
        If RsObjNagarroNo.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjNagarroNo.Close()
        RsObjNagarroNo.Open("SELECT kanbanno FROM mkt_enagaredtl WHERE kanbanno='" & strnagno & "' and Item_code='" & strItem & "'  and UNIT_CODE = '" & gstrUNITID & "' ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsObjNagarroNo.EOF Then
            ChkNagarroNo = True
        End If
        If RsObjNagarroNo.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjNagarroNo.Close()
        RsObjNagarroNo = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
    Private Function ReadPDSFile() As Boolean
        Dim strSQl As String
        Dim objCom As New SqlCommand()
        Dim currentRow As String()
        Dim oSqlCn As New SqlConnection
        Dim oSqlTran As SqlTransaction = Nothing
        ReadPDSFile = False
        Try
            oSqlCn = New SqlConnection
            oSqlCn = SqlConnectionclass.GetConnection
            oSqlTran = oSqlCn.BeginTransaction
            With objCom
                .Connection = oSqlCn
                .Transaction = oSqlTran
            End With
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFileName.Text.Trim)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters("|")
                SqlConnectionclass.ExecuteNonQuery(SqlConnectionclass.GetConnection(), CommandType.Text, "DELETE FROM TMP_PDS_FILE WHERE UNIT_CODE = '" & gstrUnitId & "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "'")
                PictureBox1.Visible = True
                While Not MyReader.EndOfData
                    Try
                        Application.DoEvents()
                        currentRow = MyReader.ReadFields()
                        If currentRow.Length < 59 Then
                            oSqlTran.Rollback()
                            objCom = Nothing
                            If oSqlCn.State = ConnectionState.Open Then oSqlCn.Close()
                            oSqlCn = Nothing
                            oSqlTran = Nothing
                            MessageBox.Show("Invalid file format !" + vbCr + "No. of columns in a file can't be less then 66 !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Function
                        End If
                        Call ChkPDSCUST_UpdFlag()

                        strSQl = "INSERT INTO TMP_PDS_FILE "
                        strSQl = strSQl + "( "
                        strSQl = strSQl + " UNIT_CODE,LOADID,SUPPLIER,SUPPLIER_PLANT,SUPPLIER_NAME,PLANT,"
                        strSQl = strSQl + " PLANT_NAME,RECEIVING_PLACE,ORDER_TYPE,PDS_NUMBER,EKB_ORDER_NO,"
                        strSQl = strSQl + " COLLECT_DATE,COLLECT_TIME,ARRIVAL_DATE,ARRIVAL_TIME,MAIN_ROUTE_GRP_CODE,"
                        strSQl = strSQl + " MAIN_ROUTE_ORDER_SEQ,SUB_ROUTE_GRP_CODE,SUB_ROUTE_ORDER_SEQ,CRS1_ROUTE,CRS1_DOCK,"
                        strSQl = strSQl + " CRS1_ARV_DATE,CRS1_ARV_TIME,CRS1_DPT_DATE,CRS1_DPT_TIME,CRS2_ROUTE,"
                        strSQl = strSQl + " CRS2_DOCK,CRS2_ARV_DATE,CRS2_ARV_TIME,CRS2_DPT_DATE,CRS2_DPT_TIME,"
                        strSQl = strSQl + " CRS3_ROUTE,CRS3_DOCK,CRS3_ARV_DATE,CRS3_ARV_TIME,CRS3_DPT_DATE,"
                        strSQl = strSQl + " CRS3_DPT_TIME,SUPPLIER_TYPE,LINE_NO,PART_NO,PART_NAME,"
                        strSQl = strSQl + " KANBAN_NO,LINE_ADDR,PACKING_SIZE,UNIT_QTY,PACK_QTY,"
                        strSQl = strSQl + " ZERO_ORDER,SORT_LANE,SHIPPING_DATE,SHIPPING_TIME,KB_PRINT_DATE_P,"
                        strSQl = strSQl + " KB_PRINT_TIME_P,KB_PRINT_DATE_L,KB_PRINT_TIME_L,REMARK,ORDER_RELEASE_DATE,"
                        strSQl = strSQl + " ORDER_RELEASE_TIME,MAIN_ROUTE_DATE,BILL_OUT_FLAG,SHIPPING_DOCK, IP_ADDRESS,SECONDARY_VENDOR_CODE,SECONDARY_PARTCODE "
                        strSQl = strSQl + ")"
                        strSQl = strSQl + " VALUES('" & gstrUnitId & "'," & Val(currentRow(0)) & ",'" & currentRow(1) & "','" & currentRow(2) & "','" & currentRow(3) & "','" & currentRow(4) & "',"
                        strSQl = strSQl + " '" & currentRow(5) & "','" & currentRow(6) & "'," & Val(currentRow(7)) & ",'" & currentRow(8) & "','" & currentRow(9) & "',"
                        strSQl = strSQl + " '" & Convert.ToDateTime(currentRow(10)).ToString("dd MMM yyyy") & "','" & currentRow(11) & "','" & Convert.ToDateTime(currentRow(12)).ToString("dd MMM yyyy") & "','" & currentRow(13) & "','" & currentRow(14) & "',"
                        '10162423
                        'strSQl = strSQl + " '" & currentRow(15) & "','" & currentRow(16) & "'," & currentRow(17) & ",'" & currentRow(18) & "','" & currentRow(19) & "',"
                        strSQl = strSQl + " '" & currentRow(15) & "','" & currentRow(16) & "','" & currentRow(17) & "','" & currentRow(18) & "','" & currentRow(19) & "',"
                        '10162423 end 
                        strSQl = strSQl + " '" & currentRow(20) & "','" & currentRow(21) & "','" & currentRow(22) & "','" & currentRow(23) & "','" & currentRow(24) & "',"
                        strSQl = strSQl + " '" & currentRow(25) & "','" & currentRow(26) & "','" & currentRow(27) & "','" & currentRow(28) & "','" & currentRow(29) & "',"
                        strSQl = strSQl + " '" & currentRow(30) & "','" & currentRow(31) & "','" & currentRow(32) & "','" & currentRow(33) & "','" & currentRow(34) & "',"
                        If SchPDSCUSTUpdFlag = False Then
                            strSQl = strSQl + " '" & currentRow(35) & "','" & currentRow(36) & "'," & Val(currentRow(37)) & ",'" & currentRow(38) & "','" & currentRow(39) & "',"
                        Else
                            strSQl = strSQl + " '" & currentRow(35) & "','" & currentRow(36) & "'," & Val(currentRow(37)) & ",'" & currentRow(67) & "','" & currentRow(39).ToString & "',"
                        End If

                        strSQl = strSQl + " '" & currentRow(40) & "','" & currentRow(41) & "'," & Val(currentRow(42)) & "," & Val(currentRow(43)) & "," & Val(currentRow(44)) & ","
                        strSQl = strSQl + " '" & currentRow(45) & "','" & currentRow(46) & "','" & currentRow(47) & "','" & currentRow(48) & "','" & currentRow(49) & "',"
                        strSQl = strSQl + " '" & currentRow(50) & "','" & currentRow(51) & "','" & currentRow(52) & "','" & currentRow(53) & "','" & currentRow(54) & "',"
                        'strSQl = strSQl + " '" & currentRow(55) & "','" & currentRow(56) & "','" & currentRow(57) & "','" & currentRow(58) & "','" & gstrIpaddressWinSck & "' )"
                        If SchPDSCUSTUpdFlag = False Then
                            strSQl = strSQl + " '" & currentRow(55) & "','" & currentRow(56) & "','" & currentRow(57) & "','" & currentRow(58) & "','" & gstrIpaddressWinSck & "','" & currentRow(66).ToString & "','" & currentRow(67).ToString & "' )"
                        Else
                            strSQl = strSQl + " '" & currentRow(55) & "','" & currentRow(56) & "','" & currentRow(57) & "','" & currentRow(58) & "','" & gstrIpaddressWinSck & "','" & currentRow(66).ToString & "','" & currentRow(38).ToString & "' )"
                        End If

                        With objCom
                            .CommandText = strSQl
                            .CommandType = CommandType.Text
                            .ExecuteNonQuery()
                        End With
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        oSqlTran.Rollback()
                        objCom = Nothing
                        If oSqlCn.State = ConnectionState.Open Then oSqlCn.Close()
                        oSqlCn = Nothing
                        oSqlTran = Nothing
                        MessageBox.Show("Line " & ex.Message & "is not valid and will be skipped.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Function
                    End Try
                End While

                With objCom
                    .Parameters.Clear()
                    .CommandTimeout = 0
                    .CommandText = "USP_UPLOAD_TKM_SCHEDULE"
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                    .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 12).Value = TxtCustomerCode.Text.Trim.ToUpper
                    .Parameters.Add("@FILE_TYPE", SqlDbType.VarChar, 50).Value = IIf(OptPDS.Checked = True, "OEM", "SPD")
                    .Parameters.Add("@FILE_PATH", SqlDbType.VarChar, 3000).Value = lblFileName.Text.Trim
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    .Parameters.Add("@USERID", SqlDbType.VarChar, 16).Value = mP_User
                    .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                    .ExecuteNonQuery()
                End With
                If objCom.Parameters(objCom.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                    PopulateScheduleLog()
                    oSqlTran.Rollback()
                    If oSqlCn.State = ConnectionState.Open Then oSqlCn.Close()
                    oSqlCn = Nothing
                    oSqlTran = Nothing
                    PictureBox1.Visible = False
                    MessageBox.Show(objCom.Parameters(objCom.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    objCom = Nothing
                    Exit Function
                End If
                oSqlTran.Commit()
                objCom = Nothing
                If oSqlCn.State = ConnectionState.Open Then oSqlCn.Close()
                oSqlCn = Nothing
                oSqlTran = Nothing
                ReadPDSFile = True
                PictureBox1.Visible = False
                MessageBox.Show("File Uploaded Successfully !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            oSqlTran.Rollback()
            objCom = Nothing
            If oSqlCn.State = ConnectionState.Open Then oSqlCn.Close()
            oSqlCn = Nothing
            oSqlTran = Nothing
            MessageBox.Show("Failed to uploaded file for following reasons:" & vbCr & ex.Message.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Finally
            PictureBox1.Visible = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Function ReadMaterialReceiptFile() As Boolean

        Dim oxlapp As Excel.Application
        Dim oxlbook As Excel.Workbook
        Dim oRange As Excel.Range
        Dim oSheet As Excel.Worksheet
        Dim strSQL As String = ""
        Dim objCom As New SqlCommand()
        Dim strReleaseNumber As String = ""
        Dim strReceiptNo As String = ""
        Dim strReceiptDate As String = ""
        Dim LngInvoiceNo As Long = 0
        Dim strItemID As String = ""
        Dim dblOrderedQty As Double = 0
        Dim dblReceivedQty As Double = 0
        Dim irow As Integer
        Dim blnHeaderFound As Boolean = False
        Dim sqlCmd As SqlCommand

        ReadMaterialReceiptFile = False

        Try
            oxlapp = (CreateObject("Excel.Application"))
            oxlbook = oxlapp.Workbooks.Open(lblFileName.Text)
            oSheet = oxlbook.Worksheets(1)
            Dim lrow As Long = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Rows.Row
            For irow = 1 To lrow
                oRange = CType(oxlbook.Worksheets(1).Cells(irow, 1), Excel.Range)
                If Not oRange.Value Is Nothing Then
                    If oRange.Value.ToString().Trim.ToUpper = "Sl No".ToUpper Then
                        blnHeaderFound = True
                        Exit For
                    End If
                End If
            Next irow

            If blnHeaderFound = False Then
                MessageBox.Show("File not in proper format !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                oxlbook = Nothing
                oRange = Nothing
                oSheet = Nothing
                oxlapp.Quit()
                oxlapp = Nothing
                Exit Function
            End If

            Dim ocells As Object = oSheet.Range("A" & irow).EntireRow.Value
            If ocells(1, 1).ToString.ToUpper.Trim <> "Sl No".ToUpper Or ocells(1, 2).ToString.ToUpper.Trim <> "Release Number".ToUpper Or ocells(1, 3).ToString.ToUpper.Trim <> "Receipt Number".ToUpper Or ocells(1, 4).ToString.ToUpper.Trim <> "Receipt Date".ToUpper Or ocells(1, 5).ToString.ToUpper.Trim <> "Invoice Number".ToUpper Or ocells(1, 6).ToString.ToUpper.Trim <> "Vendor Name".ToUpper Or ocells(1, 7).ToString.ToUpper.Trim <> "Vendor ID".ToUpper Or ocells(1, 8).ToString.ToUpper.Trim <> "Vendor Site Code".ToUpper Or ocells(1, 9).ToString.ToUpper.Trim <> "Item ID".ToUpper Or ocells(1, 10).ToString.ToUpper.Trim <> "Item Description".ToUpper Or ocells(1, 11).ToString.ToUpper.Trim <> "UOM".ToUpper Or ocells(1, 12).ToString.ToUpper.Trim <> "Ordered Quantity".ToUpper Or ocells(1, 13).ToString.ToUpper.Trim <> "Received Quantity".ToUpper Or ocells(1, 14).ToString.ToUpper.Trim <> "Difference".ToUpper Or ocells(1, 15).ToString.ToUpper.Trim <> "QTY*PRICE".ToUpper Or ocells(1, 16).ToString.ToUpper.Trim <> "Created By".ToUpper Then
                MessageBox.Show("File not in proper format !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                oxlbook = Nothing
                oRange = Nothing
                oSheet = Nothing
                oxlapp.Quit()
                oxlapp = Nothing
                Exit Function
            End If

            PictureBox1.Visible = True

            SqlConnectionclass.BeginTrans()
            SqlConnectionclass.ExecuteNonQuery("DELETE FROM SCHEDULE_UPLOAD_LOG	WHERE UNIT_CODE ='" & gstrUnitId & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'")

            While irow < lrow
                Application.DoEvents()
                irow += 1
                ocells = Nothing
                ocells = oSheet.Range("A" & irow).EntireRow.Value
                If Not ocells Is Nothing Then
                    strReleaseNumber = ""
                    strReceiptNo = ""
                    strReceiptDate = ""
                    LngInvoiceNo = 0
                    strItemID = ""
                    dblOrderedQty = 0
                    dblReceivedQty = 0
                    If Not ocells(1, 2) Is Nothing Then
                        strReleaseNumber = ocells(1, 2).ToString.ToUpper.Trim
                    End If
                    If Not ocells(1, 3) Is Nothing Then
                        strReceiptNo = ocells(1, 3).ToString.ToUpper.Trim
                    End If
                    If Not ocells(1, 4) Is Nothing Then
                        strReceiptDate = ocells(1, 4).ToString.ToUpper.Trim
                    End If
                    If Not ocells(1, 5) Is Nothing Then
                        LngInvoiceNo = Val(ocells(1, 5).ToString.ToUpper.Trim)
                    End If
                    If Not ocells(1, 9) Is Nothing Then
                        strItemID = Replace(Replace(ocells(1, 9).ToString.ToUpper.Trim, "=", ""), Chr(34), "")
                    End If
                    If Not ocells(1, 12) Is Nothing Then
                        dblOrderedQty = ocells(1, 12).ToString.ToUpper.Trim
                    End If
                    If Not ocells(1, 13) Is Nothing Then
                        dblReceivedQty = ocells(1, 13).ToString.ToUpper.Trim
                    End If
                    If strReleaseNumber <> "" Then
                        With objCom
                            .Parameters.Clear()
                            .CommandText = "USP_MATERIAL_RECEIPT_AGAINST_PDS"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                            .Parameters.Add("@RELEASENUMBER", SqlDbType.VarChar, 50).Value = strReleaseNumber
                            .Parameters.Add("@RECEIPTNO", SqlDbType.VarChar, 50).Value = strReceiptNo
                            .Parameters.Add("@RECEIPTDATE", SqlDbType.DateTime).Value = VB6.Format(strReceiptDate, "dd MMM yyyy")
                            .Parameters.Add("@INVOICENO", SqlDbType.Money).Value = LngInvoiceNo
                            .Parameters.Add("@ITEMID", SqlDbType.VarChar, 50).Value = strItemID
                            .Parameters.Add("@ORDQTY", SqlDbType.Money).Value = dblOrderedQty
                            .Parameters.Add("@RECQTY", SqlDbType.Money).Value = dblReceivedQty
                            .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                            .Parameters.Add("@USERID", SqlDbType.VarChar, 16).Value = mP_User
                            .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                            SqlConnectionclass.ExecuteNonQuery(objCom)
                        End With
                        If objCom.Parameters(objCom.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                            SqlConnectionclass.RollbackTran()
                            objCom = Nothing
                            PictureBox1.Visible = False
                            MessageBox.Show(objCom.Parameters(objCom.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Function
                        End If
                    End If
                End If
            End While

            SqlConnectionclass.CommitTran()
            ReadMaterialReceiptFile = True
            objCom = Nothing
            oxlbook.Close()
            oxlbook = Nothing
            oRange = Nothing
            oSheet = Nothing
            oxlapp.Quit()
            oxlapp = Nothing

            If Not IsNothing(objCom) Then objCom.Dispose()
            PopulateScheduleLog()
            PictureBox1.Visible = False
            MessageBox.Show("File Uploaded Successfully !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)

            'changes done by Abhinav against issue ID 10715594 
            objCom = New SqlCommand
            With objCom
                .CommandText = "USP_PENDING_PDS_AUTOMAILER"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 10).Value = gstrIpaddressWinSck
                .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                SqlConnectionclass.ExecuteNonQuery(objCom)
            End With

            If objCom.Parameters(objCom.Parameters.Count - 1).Value.ToString.Trim.Length <> 0 Then
                MessageBox.Show("Error while sending the Mail !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Function
            End If
            'changes exterminates here (10715594)

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Not IsNothing(objCom) Then objCom.Dispose()
            PictureBox1.Visible = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            oxlbook = Nothing
            oRange = Nothing
            oSheet = Nothing
            oxlapp = Nothing
        End Try
    End Function
#End Region
#Region "Form Events"
    Private Sub FRMMKTTRN0072_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try
            mdifrmMain.CheckFormName = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0072_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0072_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            mstrIPAddress = gstrIpaddressWinSck
            Call FitToClient(Me, fraMain, ctlFormHeader1, fraButton, 300)
            fraUpload.Visible = True
            frabuttonCancellation.Visible = False
            Call EnableControls(True, Me, True)
            MdiParent = prjMPower.mdifrmMain
            gblnCancelUnload = False : gblnFormAddEdit = False
            mintFormtag = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader1.Tag)
            ctlFormHeader1.HeaderString = "Sales Schedule and Material Receipt Uploading"
            Me.Text = "Sales Schedule and Material Receipt Uploading"
            TxtCustomerCode.ForeColor = Color.Crimson
            PictureBox1.Visible = False
            sspr.Visible = False
            OptPDS.Checked = True
            fraUpload.Visible = True
            frabuttonCancellation.Visible = False
            Call ChkVBSchUpdFlag()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0072_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
                Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0072_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Escape
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        TxtCustomerCode.Text = ""
                        lblCustDesc.Text = ""
                        lblFileName.Text = ""
                    Else
                        Me.ActiveControl.Focus()
                    End If
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            e.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        End Try
    End Sub
    Private Sub FRMMKTTRN0072_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
            mdifrmMain.RemoveFormNameFromWindowList = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
#End Region
#Region "Other Controls Events"
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub TxtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    If CmdBrowse.Enabled = True Then CmdBrowse.Focus()
                Case 39, 34, 96
                    KeyAscii = 0
                Case 13
                    SendKeys.Send("{TAB}")
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        End Try
    End Sub
    Private Sub TxtCustomerCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                If CmdHlpCustomer.Enabled Then Call CmdHlpCustomer_Click(CmdHlpCustomer, New System.EventArgs())
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub TxtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtCustomerCode.Validating
        Dim blnCancel As Boolean = False
        Try
            If TxtCustomerCode.Text.Trim.Length = 0 Then
                lblCustDesc.Text = ""
                blnCancel = False
            Else
                'Changes against 10737738 
                If SchUpdFlag = True Then
                    blnCancel = Not IsExists("SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUnitId & "' AND CUSTOMER_CODE='" + TxtCustomerCode.Text.Trim + "' and SCH_UPLOAD_CODE ='PDS' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106))) ")
                Else
                    blnCancel = Not IsExists("SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUnitId & "' AND CUSTOMER_CODE='" + TxtCustomerCode.Text.Trim + "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106))) ")
                End If
                If blnCancel = True Then
                    MessageBox.Show("Customer code doesn't exists !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    TxtCustomerCode.Text = ""
                    lblCustDesc.Text = ""
                End If
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            e.Cancel = blnCancel
        End Try
    End Sub
    Private Sub CmdHlpCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdHlpCustomer.Click
        Dim strsql() As String
        Try
            'Changes against 10737738 
            If OptSeqPDSFile.Checked Then
                strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE ='" & gstrUnitId & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106))) AND ISNULL(PDS_UPLOAD_SECONDARYPARTCODE_FLAG, 0)=1 ORDER BY CUSTOMER_CODE", "Customer Help", 1)
            ElseIf optPDSMaruti.Checked Then
                strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE ='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106))) AND ISNULL(PDS_UPLOAD_SECONDARYPARTCODE_FLAG, 0)=0 AND ISNULL(PDSNagaare,0)=1 ORDER BY CUSTOMER_CODE", "Customer Help", 1)
            Else
                If SchUpdFlag = True Then
                    strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE ='" & gstrUNITID & "' and SCH_UPLOAD_CODE ='PDS' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106))) ORDER BY CUSTOMER_CODE", "Customer Help", 1)
                Else
                    strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE ='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106))) ORDER BY CUSTOMER_CODE", "Customer Help", 1)
                End If
            End If
            If Not (UBound(strsql) <= 0) Then
                If Not (UBound(strsql) = 0) Then
                    If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                        MsgBox("No Customer To Display.", MsgBoxStyle.Information, ResolveResString(100))
                        TxtCustomerCode.Text = ""
                        lblCustDesc.Text = ""
                        Exit Sub
                    Else
                        TxtCustomerCode.Text = strsql(0).Trim
                        lblCustDesc.Text = strsql(1).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub OptPDS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptPDS.CheckedChanged, optSPD.CheckedChanged, OptMaterialReceipt.CheckedChanged, OptSeqPDSFile.CheckedChanged, optPDSMaruti.CheckedChanged
        OptPDS.ForeColor = lblCustDesc.BackColor
        optSPD.ForeColor = lblCustDesc.BackColor
        OptMaterialReceipt.ForeColor = lblCustDesc.BackColor
        OptSeqPDSFile.ForeColor = lblCustDesc.BackColor
        optPDSMaruti.ForeColor = lblCustDesc.BackColor
        SetSpreadProperty()
        sspr.Visible = False
        lblFileName.Text = String.Empty
        If OptPDS.Checked = True Then
            OptPDS.ForeColor = Color.DarkGreen
        ElseIf optSPD.Checked = True Then
            optSPD.ForeColor = Color.DarkGreen
        ElseIf OptMaterialReceipt.Checked = True Then
            OptMaterialReceipt.ForeColor = Color.DarkGreen
        ElseIf OptSeqPDSFile.Checked = True Then
            OptSeqPDSFile.ForeColor = Color.DarkGreen
        ElseIf optPDSMaruti.Checked = True Then
            optPDSMaruti.ForeColor = Color.DarkGreen
        End If
    End Sub
    Private Sub TabMain_Selected(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabMain.Selected
        Try
            If e.TabPage.Equals(TabUploading) Then
                fraUpload.Visible = True
                frabuttonCancellation.Visible = False
            Else
                fraUpload.Visible = False
                frabuttonCancellation.Visible = True
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Function IsExists(ByVal strSql As String) As Boolean
        Dim objconn As SqlConnection = Nothing
        Dim objDR As SqlDataReader
        Dim objCommand As New SqlCommand()
        IsExists = False
        If strSql.Trim.Length = 0 Then Exit Function
        Try
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = strSql
            objCommand.CommandType = CommandType.Text
            objDR = objCommand.ExecuteReader()
            If objDR.HasRows = True Then
                IsExists = True
            End If
            objDR.Close()
            objDR = Nothing
            objCommand = Nothing
            objconn.Close()
            objconn = Nothing
        Catch ex As Exception
            objDR = Nothing
            objCommand = Nothing
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objconn = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
#End Region
#Region "Command Controls"
    Private Sub CmdBrowse_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdBrowse.Click
        Dim oFDialog As New OpenFileDialog()
        Try
            sspr.Visible = False
            oFDialog.Filter = IIf(OptMaterialReceipt.Checked = False, "Text files (*.txt)|*.txt", "Excel files (*.xls)|*.xls")
            oFDialog.FilterIndex = 3
            oFDialog.RestoreDirectory = True
            oFDialog.Title = "Select a file to upload"
            If oFDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                lblFileName.Tag = oFDialog.SafeFileName
                lblFileName.Text = oFDialog.FileName
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub BtnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Try
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub CmdUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdUpload.Click
        Try
            sspr.Visible = False
            If OptMaterialReceipt.Checked = False And OptPDS.Checked = False And optSPD.Checked = False And OptSeqPDSFile.Checked = False And optPDSMaruti.Checked = False Then
                MessageBox.Show("Select a File Type option to be uploaded !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If TxtCustomerCode.Text.Trim.Length = 0 Then
                MessageBox.Show("Select a customer !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                If TxtCustomerCode.Enabled = True Then TxtCustomerCode.Focus()
                Exit Sub
            End If
            If lblFileName.Text.Trim.Length = 0 Then
                MessageBox.Show("Select a file to upload !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                If CmdBrowse.Enabled = True Then CmdBrowse.Focus()
                Exit Sub
            End If
            If OptPDS.Checked = True Or optSPD.Checked = True Then
                If ReadPDSFile() = False Then
                    Exit Sub
                End If
            ElseIf OptSeqPDSFile.Checked = True Then
                If ReadSequencePDSFile() = False Then
                    Exit Sub
                End If
            ElseIf optPDSMaruti.Checked Then
                If ReadMarutiPDSFile() = False Then
                    Exit Sub
                End If
            Else
                If ReadMaterialReceiptFile() = False Then
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            PictureBox1.Visible = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub CmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdClose.Click, BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
#End Region

    Private Function ReadSequencePDSFile() As Boolean
        Dim currentRow As String()
        Dim dtSeqPDS As New DataTable
        Dim sqlCmd As New SqlCommand
        Try
            ReadSequencePDSFile = False
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFileName.Text.Trim)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters("|")
                PictureBox1.Visible = True

                dtSeqPDS.Columns.Add("LOADID", GetType(System.Int32))
                dtSeqPDS.Columns.Add("SUPPLIER", GetType(System.String))
                dtSeqPDS.Columns.Add("SUPPLIER_PLANT", GetType(System.String))
                dtSeqPDS.Columns.Add("SUPPLIER_NAME", GetType(System.String))
                dtSeqPDS.Columns.Add("PLANT", GetType(System.String))
                dtSeqPDS.Columns.Add("PLANT_NAME", GetType(System.String))
                dtSeqPDS.Columns.Add("RECEIVING_PLACE", GetType(System.String))
                dtSeqPDS.Columns.Add("ORDER_TYPE", GetType(System.Int32))
                dtSeqPDS.Columns.Add("PDS_NUMBER", GetType(System.String))
                dtSeqPDS.Columns.Add("EKB_ORDER_NO", GetType(System.String))
                dtSeqPDS.Columns.Add("COLLECT_DATE", GetType(System.DateTime))
                dtSeqPDS.Columns.Add("COLLECT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("ARRIVAL_DATE", GetType(System.DateTime))
                dtSeqPDS.Columns.Add("ARRIVAL_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("MAIN_ROUTE_GRP_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("MAIN_ROUTE_ORDER_SEQ", GetType(System.String))
                dtSeqPDS.Columns.Add("SUB_ROUTE_GRP_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("SUB_ROUTE_ORDER_SEQ", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_ARV_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_ARV_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_DPT_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_DPT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_ARV_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_ARV_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_DPT_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_DPT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_ARV_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_ARV_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_DPT_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_DPT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("SUPPLIER_TYPE", GetType(System.String))
                dtSeqPDS.Columns.Add("LINE_NO", GetType(System.Int32))
                dtSeqPDS.Columns.Add("PART_NO", GetType(System.String))
                dtSeqPDS.Columns.Add("PART_NAME", GetType(System.String))
                dtSeqPDS.Columns.Add("KANBAN_NO", GetType(System.String))
                dtSeqPDS.Columns.Add("SEQ_NO", GetType(System.Int64))
                dtSeqPDS.Columns.Add("PACKING_SIZE", GetType(System.Int32))
                dtSeqPDS.Columns.Add("UNIT_QTY", GetType(System.Int32))
                dtSeqPDS.Columns.Add("PACK_QTY", GetType(System.Int32))
                dtSeqPDS.Columns.Add("ZERO_ORDER", GetType(System.String))
                dtSeqPDS.Columns.Add("SORT_LANE", GetType(System.String))
                dtSeqPDS.Columns.Add("SHIPPING_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("SHIPPING_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_DATE_P", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_TIME_P", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_DATE_L", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_TIME_L", GetType(System.String))
                dtSeqPDS.Columns.Add("REMARK", GetType(System.String))
                dtSeqPDS.Columns.Add("ORDER_RELEASE_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("ORDER_RELEASE_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("MAIN_ROUTE_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("BILL_OUT_FLAG", GetType(System.String))
                dtSeqPDS.Columns.Add("SHIPPING_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("PACKING_TYPE", GetType(System.String))
                dtSeqPDS.Columns.Add("KANBAN_ORIENTATION", GetType(System.String))
                dtSeqPDS.Columns.Add("DOLLY_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("PICKER_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CYCLE_SERIAL", GetType(System.String))
                dtSeqPDS.Columns.Add("PACKING_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("TKM_LINE_ADDRESS", GetType(System.String))
                dtSeqPDS.Columns.Add("BPA_NUMBER", GetType(System.String))
                dtSeqPDS.Columns.Add("SECONDARY_VENDOR_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("SECONDARY_PARTCODE", GetType(System.String))

                Dim drSeqPDS As DataRow
                While Not MyReader.EndOfData
                    Try
                        Application.DoEvents()
                        currentRow = MyReader.ReadFields()
                        If currentRow.Length < 68 Then
                            PictureBox1.Visible = False
                            MessageBox.Show("Invalid file format !" + vbCr + "No. of columns in a file can't be less then 68 !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Function
                        End If

                        drSeqPDS = dtSeqPDS.NewRow()
                        drSeqPDS("LOADID") = Val(currentRow(0))
                        drSeqPDS("SUPPLIER") = currentRow(1)
                        drSeqPDS("SUPPLIER_PLANT") = currentRow(2)
                        drSeqPDS("SUPPLIER_NAME") = currentRow(3)
                        drSeqPDS("PLANT") = currentRow(4)
                        drSeqPDS("PLANT_NAME") = currentRow(5)
                        drSeqPDS("RECEIVING_PLACE") = currentRow(6)
                        drSeqPDS("ORDER_TYPE") = Val(currentRow(7))
                        drSeqPDS("PDS_NUMBER") = currentRow(8)
                        drSeqPDS("EKB_ORDER_NO") = currentRow(9)
                        drSeqPDS("COLLECT_DATE") = Convert.ToDateTime(currentRow(10)).ToString("dd MMM yyyy")
                        drSeqPDS("COLLECT_TIME") = currentRow(11)
                        drSeqPDS("ARRIVAL_DATE") = Convert.ToDateTime(currentRow(12)).ToString("dd MMM yyyy")
                        drSeqPDS("ARRIVAL_TIME") = currentRow(13)
                        drSeqPDS("MAIN_ROUTE_GRP_CODE") = currentRow(14)
                        drSeqPDS("MAIN_ROUTE_ORDER_SEQ") = currentRow(15)
                        drSeqPDS("SUB_ROUTE_GRP_CODE") = currentRow(16)
                        drSeqPDS("SUB_ROUTE_ORDER_SEQ") = currentRow(17)
                        drSeqPDS("CRS1_ROUTE") = currentRow(18)
                        drSeqPDS("CRS1_DOCK") = currentRow(19)
                        drSeqPDS("CRS1_ARV_DATE") = currentRow(20)
                        drSeqPDS("CRS1_ARV_TIME") = currentRow(21)
                        drSeqPDS("CRS1_DPT_DATE") = currentRow(22)
                        drSeqPDS("CRS1_DPT_TIME") = currentRow(23)
                        drSeqPDS("CRS2_ROUTE") = currentRow(24)
                        drSeqPDS("CRS2_DOCK") = currentRow(25)
                        drSeqPDS("CRS2_ARV_DATE") = currentRow(26)
                        drSeqPDS("CRS2_ARV_TIME") = currentRow(27)
                        drSeqPDS("CRS2_DPT_DATE") = currentRow(28)
                        drSeqPDS("CRS2_DPT_TIME") = currentRow(29)
                        drSeqPDS("CRS3_ROUTE") = currentRow(30)
                        drSeqPDS("CRS3_DOCK") = currentRow(31)
                        drSeqPDS("CRS3_ARV_DATE") = currentRow(32)
                        drSeqPDS("CRS3_ARV_TIME") = currentRow(33)
                        drSeqPDS("CRS3_DPT_DATE") = currentRow(34)
                        drSeqPDS("CRS3_DPT_TIME") = currentRow(35)
                        drSeqPDS("SUPPLIER_TYPE") = currentRow(36)
                        drSeqPDS("LINE_NO") = Val(currentRow(37))
                        drSeqPDS("PART_NO") = currentRow(38)
                        drSeqPDS("PART_NAME") = currentRow(39)
                        drSeqPDS("KANBAN_NO") = currentRow(40)
                        drSeqPDS("SEQ_NO") = Val(currentRow(41))
                        drSeqPDS("PACKING_SIZE") = Val(currentRow(42))
                        drSeqPDS("UNIT_QTY") = Val(currentRow(43))
                        drSeqPDS("PACK_QTY") = Val(currentRow(44))
                        drSeqPDS("ZERO_ORDER") = currentRow(45)
                        drSeqPDS("SORT_LANE") = currentRow(46)
                        drSeqPDS("SHIPPING_DATE") = currentRow(47)
                        drSeqPDS("SHIPPING_TIME") = currentRow(48)
                        drSeqPDS("KB_PRINT_DATE_P") = currentRow(49)
                        drSeqPDS("KB_PRINT_TIME_P") = currentRow(50)
                        drSeqPDS("KB_PRINT_DATE_L") = currentRow(51)
                        drSeqPDS("KB_PRINT_TIME_L") = currentRow(52)
                        drSeqPDS("REMARK") = currentRow(53)
                        drSeqPDS("ORDER_RELEASE_DATE") = currentRow(54)
                        drSeqPDS("ORDER_RELEASE_TIME") = ""
                        drSeqPDS("MAIN_ROUTE_DATE") = currentRow(55)
                        drSeqPDS("BILL_OUT_FLAG") = currentRow(56)
                        drSeqPDS("SHIPPING_DOCK") = currentRow(57)
                        drSeqPDS("PACKING_TYPE") = currentRow(58)
                        drSeqPDS("KANBAN_ORIENTATION") = currentRow(59)
                        drSeqPDS("DOLLY_CODE") = currentRow(60)
                        drSeqPDS("PICKER_ROUTE") = currentRow(61)
                        drSeqPDS("CYCLE_SERIAL") = currentRow(62)
                        drSeqPDS("PACKING_CODE") = currentRow(63)
                        drSeqPDS("TKM_LINE_ADDRESS") = currentRow(64)
                        drSeqPDS("BPA_NUMBER") = currentRow(65)
                        drSeqPDS("SECONDARY_VENDOR_CODE") = currentRow(66)
                        drSeqPDS("SECONDARY_PARTCODE") = currentRow(67)

                        dtSeqPDS.Rows.Add(drSeqPDS)
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MessageBox.Show("Line " & ex.Message & "is not valid and will be skipped.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Function
                    End Try
                End While

                If dtSeqPDS IsNot Nothing AndAlso dtSeqPDS.Rows.Count > 0 Then
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandTimeout = 600 ' 10 Minute
                        .CommandText = "USP_UPLOAD_PDS_SEQUENCING_SCHEDULE"
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                        .Parameters.AddWithValue("@CUSTOMER_CODE", TxtCustomerCode.Text.Trim())
                        .Parameters.AddWithValue("@FILE_TYPE", "OEM SEQ")
                        .Parameters.AddWithValue("@FILE_PATH", lblFileName.Text.Trim)
                        .Parameters.AddWithValue("@USER_ID", mP_User)
                        .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                        .Parameters.AddWithValue("@UDT_PDS_SEQUENCING_FILE", dtSeqPDS)
                        .Parameters.AddWithValue("@ACTION", "UPLOAD_MANUALLY")
                        .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                        If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                            PopulateScheduleLog()
                            PictureBox1.Visible = False
                            MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Function
                        Else
                            ReadSequencePDSFile = True
                            PictureBox1.Visible = False
                            MessageBox.Show("File Uploaded Successfully !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End With
                End If
            End Using
        Catch ex As Exception
            RaiseException(ex)
        Finally
            PictureBox1.Visible = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If dtSeqPDS IsNot Nothing Then
                dtSeqPDS.Dispose()
            End If
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
            End If
        End Try
    End Function
End Class