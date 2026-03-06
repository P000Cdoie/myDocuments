Option Strict Off
Option Explicit On
Friend Class FRMMKTTRN0073
    Inherits System.Windows.Forms.Form
    Dim Conn_For_Trigger As New ADODB.Connection
    Dim mP_ServerName As String
    Dim R_VLog_id As String
    Dim R_VPword As String
    Dim mP_DatabaseName As String
    Dim IsTrans As Boolean = False
    '***********************************************************************
    'Revision History
    'Modified By        :   Geetanjali Aggrawal
    'Modified On        :   16 May 2014
    'Purpose            :   10596054 - if records already exists then first delete then insert.
    '************************************************************************
    Private Sub cmdFileHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFileHelp.Click
        Dim gobjError As Object
        On Error GoTo ErrHandler
        CommanDlogOpen.InitialDirectory = gstrLocalCDrive
        CommanDlogOpen.Filter = "Microsoft Excel File (*.xls)|*.xls"
        CommanDlogOpen.ShowDialog()
        Me.txtFileName.Text = CommanDlogOpen.FileName
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, Conn_For_Trigger)
    End Sub

    Public Function triggerdataporting(ByVal customercode As String, ByVal dbf_filepath As String) As Object
        Dim Rsdoc As Object
        On Error GoTo Upload_ERR
        Dim cnnDB As ADODB.Connection
        Dim strDBPath As String
        Dim rstResultSet As New ADODB.Recordset
        Dim cmdDataInsert As New ADODB.Command
        Dim strsql As String
        Dim lngCurRecPos As Integer
        Dim StrSeqNo As String
        Dim Doc_Dt As Date
        Dim strTime As String
        Dim StrVINNo As String
        Dim StrCat_Code As String
        Dim StrModel_Desc As String
        Dim StrBody_color As String
        Dim StrModel_No As String
        Dim StrSeq_Desc As String
        Dim strRemark As String
        Dim StrTrackPoint As String
        Dim rsdb As ADODB.Recordset
        rsdb = New ADODB.Recordset
        'Open DBF file and read data
        strDBPath = dbf_filepath
        Dim AlreadyUpload As Boolean = True
        Dim TimeFomat As Boolean = True
        Dim st As String
        IsTrans = False

        'mP_ServerName = "172.29.29.11\sql2005"
        'R_VLog_id = "user"
        'R_VPword = "catseye"
        'mP_DatabaseName = "xfinii_2005"

        mP_ServerName = gstrCONNECTIONSERVER ' "172.29.96.25"
        R_VLog_id = "user"
        R_VPword = "catseye"
        mP_DatabaseName = gstrCONNECTIONDATABASE ' "Test_M1EMPOWER"

        If Conn_For_Trigger.State = ADODB.ObjectStateEnum.adStateOpen Then
            Conn_For_Trigger.Close()
        End If

        With Conn_For_Trigger
            .ConnectionString = "Provider=sqloledb;Data Source=" & Trim(mP_ServerName) & ";Initial Catalog=" & Trim(mP_DatabaseName) & ";User Id=" & Trim(R_VLog_id) & ";Password=" & Trim(R_VPword) & ";"
            .CursorLocation = ADODB.CursorLocationEnum.adUseServer
            .Open()
            Conn_For_Trigger.CommandTimeout = 0
            .BeginTrans()
            IsTrans = True
            .Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End With
        'Query the file
        strsql = "SELECT * FROM [sheet1$] "

        cnnDB = New ADODB.Connection

        With cnnDB
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Properties("Extended Properties").Value = "Excel 8.0"
            .Open(strDBPath)
        End With

        'Reading from File
        With rstResultSet
            .let_ActiveConnection(cnnDB)
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .CursorType = ADODB.CursorTypeEnum.adOpenStatic
            .LockType = ADODB.LockTypeEnum.adLockReadOnly
            .Open(strsql)
        End With
        cnnDB = Nothing
        '--Go for Insertion

        While Not rstResultSet.EOF
            '--For Line Number as well as event generation
            lngCurRecPos = lngCurRecPos + 1
            If lngCurRecPos = 1 Then
                rstResultSet.MoveFirst()
            End If
            If IsDBNull(rstResultSet.Collect(0)) = True Then Exit Function
            StrSeqNo = Trim(rstResultSet.Collect(0))
            Doc_Dt = rstResultSet.Collect(1)
            strTime = Trim(rstResultSet.Collect(2))
            Dim dt As DateTime
            If DateTime.TryParse(strTime, dt) Then
                strTime = dt.ToString("HH:mm:ss")
            Else
                TimeFomat = False

            End If
            StrVINNo = Trim(rstResultSet.Collect(3))
            StrCat_Code = Trim(rstResultSet.Collect(4))
            StrBody_color = Trim(rstResultSet.Collect(6))
            StrModel_No = Trim(rstResultSet.Collect(7))
            StrSeq_Desc = Replace(Mid(Trim(rstResultSet.Collect(8)), 1, 40), "'", "")
            strRemark = Replace(Mid(Trim(rstResultSet.Collect(9)), 1, 40), "'", "")
            StrTrackPoint = Trim(rstResultSet.Collect(10))
            strsql = "selecT Plant_c From plant_mst where UNIT_CODE = '" & gstrUNITID & "'"
            Rsdoc = Conn_For_Trigger.Execute(strsql)
            If Rsdoc.EOF <> True Then
                strsql = ""
                '--Insert Statement
                strsql = "Select count(*) RecordCnt from trigger_uploading (NOLOCK) where UNIT_CODE = '" & gstrUNITID & _
                         "' and Vin_No = '" & StrVINNo & "' and trig_loc = '" & Trim(Me.cmblist.Text) & _
                         "' and convert(datetime,Doc_Date,103)=convert(datetime,'" & Doc_Dt & "',103)"
                rsdb = Conn_For_Trigger.Execute(strsql)
                '10596054 - Added by Geetanjali
                If rsdb.Fields("RecordCnt").Value = 0 Then
                    rsdb.Close()
                    'UPGRADE_NOTE: Object rsdb may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    rsdb = Nothing
                    strsql = ""
                    strsql = "INSERT INTO trigger_uploading"
                    strsql = strsql & "(UNIT_CODE,trig_loc,Cust_Code,Seq_No,Doc_Date,Doc_Time,Vin_No,Cat_Code,Body_Color,Model_No,Seq_Desc,Remark,Tracking_Point,Ent_Dt,Ent_Uid,Upd_Dt,Upd_Uid) "
                    strsql = strsql & "VALUES ('" & gstrUNITID & "','" & Trim(Me.cmblist.Text) & "','" & Trim(txtcustname.Text) & "', '" & StrSeqNo & "', '" & VB6.Format(Doc_Dt, "dd mmm yyyy") & "', convert(varchar(8),'" & strTime & "'), '" & StrVINNo & "', '" & StrCat_Code & "', '" & StrBody_color & "', '" & StrModel_No & "', '" & StrSeq_Desc & "', '" & strRemark & "', '" & StrTrackPoint & "',getdate(),'" & 0 & "',getdate(),'" & 0 & "' )"
                Else
                    strsql = "delete from trigger_uploading where UNIT_CODE = '" & gstrUNITID & _
                         "' and Vin_No = '" & StrVINNo & "' and trig_loc = '" & Trim(Me.cmblist.Text) & _
                         "' and convert(datetime,Doc_Date,103)=convert(datetime,'" & Doc_Dt & "',103) "
                    strsql = strsql & "INSERT INTO trigger_uploading"
                    strsql = strsql & "(UNIT_CODE,trig_loc,Cust_Code,Seq_No,Doc_Date,Doc_Time,Vin_No,Cat_Code,Body_Color,Model_No" & _
                            " ,Seq_Desc,Remark,Tracking_Point,Ent_Dt,Ent_Uid,Upd_Dt,Upd_Uid) "
                    strsql = strsql & "VALUES ('" & gstrUNITID & "','" & Trim(Me.cmblist.Text) & "','" & Trim(txtcustname.Text) & "', '" & _
                            StrSeqNo & "', '" & VB6.Format(Doc_Dt, "dd mmm yyyy") & "', convert(varchar(8),'" & strTime & "'), '" & StrVINNo & "', '" & StrCat_Code & _
                            "', '" & StrBody_color & "', '" & StrModel_No & "', '" & StrSeq_Desc & "', '" & strRemark & "', '" & StrTrackPoint & _
                            "',getdate(),'" & 0 & "',getdate(),'" & 0 & "' )"
                End If
                Conn_For_Trigger.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Conn_For_Trigger.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                AlreadyUpload = False
                rstResultSet.MoveNext()
            Else
                MsgBox("Please update the Category Code and Body Color on Customer category master for[" & StrCat_Code & "] - " & StrBody_color, MsgBoxStyle.DefaultButton1, ResolveResString(100))
                If IsTrans = True Then
                    Conn_For_Trigger.RollbackTrans()
                    IsTrans = False
                End If
                Conn_For_Trigger.Close()
                Exit Function
            End If
        End While
        rstResultSet = Nothing

        If AlreadyUpload = True Then
            If IsTrans = True Then
                Conn_For_Trigger.RollbackTrans()
                IsTrans = False
            End If
            MsgBox("The Same Data Already uploaded in the System", MsgBoxStyle.DefaultButton1, ResolveResString(100))
        End If
        Exit Function
Upload_ERR:
        If IsTrans = True Then
            Conn_For_Trigger.RollbackTrans()
            IsTrans = False
        End If

        If Err.Number = -2147467259 Then
            MsgBox(" Upload the Specified format file", MsgBoxStyle.DefaultButton1, ResolveResString(100))
        Else
            MsgBox(Err.Number & " " & Err.Description, MsgBoxStyle.DefaultButton1, ResolveResString(100))
        End If
    End Function

    Private Sub Form2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, FrmMain, ctlFormHeader1, PicBox, 500) 'To fit th
            Call FillLabelFromResFile(Me)
            Me.MdiParent = mdifrmMain

            'conmain = New ADODB.Connection
            'If DBConnection(R_VLog_id, R_VPword) = False Then
            '    MsgBox("Connection Failed", MsgBoxStyle.Critical)
            '    End
            'End If

            Conn_For_Trigger.CommandTimeout = 0
            cmblist.Items.Add(("UB"))
            cmblist.Items.Add(("MK"))
            cmblist.Items.Add(("DL"))
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub lblMsg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMsg.Click

    End Sub

    Private Sub CmdUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdUpload.Click
        Dim Trigger As Object
        Try
            Dim m_strcustcode As String
            If cmblist.Text = "" Then
                MsgBox("Please Select Trigger Location", MsgBoxStyle.OkOnly)
                cmblist.Focus()
                Exit Sub
            End If

            If txtFileName.Text = "" Then
                MsgBox("Please select file", MsgBoxStyle.OkOnly)
                txtFileName.Focus()
                Exit Sub
            End If
            lblMsg.Text = "Uploading Data ..."
            m_strcustcode = txtcustname.Text
            If triggerdataporting(m_strcustcode, Trim(Me.txtFileName.Text)) = True Then
                MsgBox("upload failed")
            Else
                lblMsg.Text = ""
                If IsTrans = True Then
                    Conn_For_Trigger.CommitTrans()
                    IsTrans = False
                    MsgBox("Data Upload Sucessfully", MsgBoxStyle.OkOnly, Trigger)
                End If
            End If
            Exit Sub
        Catch Ex As Exception
            lblMsg.Text = ""
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
        End Try
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    Private Sub Command1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command1.Click
        FRMMKTTRN0073A.ShowDialog()
    End Sub

    Private Sub FrmMain_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMain.Enter

    End Sub
End Class