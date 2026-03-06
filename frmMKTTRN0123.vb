Imports System
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine


Public Class FRMMKTTRN0123
    '========================================================================================
    'COPYRIGHT          :   MOTHERSONSUMI INFOTECH & DESIGN LTD.
    'AUTHOR             :   GEETANJALI AGGARWAL
    'CREATION DATE      :   08 Jun 2015
    'DESCRIPTION        :   10798002 - freight Managment in eMPro
    '-----------------------------------------------------------------------------------------------
    'REVISION HISTORY   -   10904085 — FW Trip Generation issue mate tapukara
    'REVISION DATE      -   28 Sep 2015
    'REVISED BY         -   Geetanjali Aggrawal
    '-----------------------------------------------------------------------------------------------
    'REVISION HISTORY   -   10916489 — FW: Trip generation issue(Repeated data selection of Document no.)
    'REVISION DATE      -   23 Oct 2015
    'REVISED BY         -   Geetanjali Aggrawal
    '-----------------------------------------------------------------------------------------------
    'REVISION HISTORY   -   10949588 — Customer Own Vehicle Functionality Addition
    'REVISION DATE      -   11 Jan 2016
    'REVISED BY         -   Geetanjali Aggrawal
    '		
    '-----------------------------------------------------------------------------------------------
    'REVISION HISTORY   -   10976251 — Trip Generation option very slow
    'REVISION DATE      -   02 Feb 2016
    'REVISED BY         -   Geetanjali Aggrawal
    '=========================================================================================================
    'REVISION HISTORY   -   101053952 — Not able to delete wrong trip MATE Tapukara.
    'REVISION DATE      -   30 May 2016
    'REVISED BY         -   Parveen Kumar

    'Modified BY         - Ashish sharma
    'Modified Date       - 10 AUG 2017
    'Issue Id            - 101335650
    'Description         - Freight Management Enhancement
    '=========================================================================================================
    'Modified By        - SWATI PAREEK
    'Created Date       - 12 june 2021
    'description        - TT#1035570 — ADHOC PHASE -2 (add a view adhoc doc button)
    '=========================================================================================================

#Region "Global variables"
    Dim dtDocDtl As DataTable
    Dim REPDOC As ReportDocument
    Dim REPVIEWER As eMProCrystalReportViewer
    Const _contractTypeTripBased As String = "TRIP"
    Const _contractTypeMonthlyBased As String = "MONTHLY"
    Const _contractTypeCourierBased As String = "SIZE_WEIGHT"
    Const _contractTypeStaffBased As String = "STAFF"
    Dim mblnAllowTransporterfromMaster As Boolean = False
    Dim blnIsAdhoc As Boolean

    Enum Enum_DocSel
        Col_DocType = 1
        Col_DocNo = 2
        Col_DocDt = 3
        Col_CustCode = 4
        Col_VendorCode = 5
        Col_FromLocation = 6
        Col_VendorWhCode = 7
        Col_Qty = 8
        Col_DocValue = 9
    End Enum

#End Region

#Region "Form Events"
    Private Sub FRMMKTTRN0123_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, GrpMain, ctlFormHeader, GrpCmdBtn, 500)
            Me.MdiParent = mdifrmMain
            txtScan.Text = "56|V1130414|Om Logistics Ltd|VC0072|10|SML"
            InitializeForm(1)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region


    Private Sub cmdGrp_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdGrp.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    InitializeForm(2)
                   
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                    txtScan.Text = ""
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If SaveData() Then
                        InitializeForm(3)
                        cmdGrp.Revert()
                        cmdGrp.Top = 14
                        cmdGrp.Left = 53
                        cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                        cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                        cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                        cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                        cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                        cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True '' At save time for Delete enable
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    InitializeForm(1)
                    cmdGrp.Revert()
                    cmdGrp.Top = 14
                    cmdGrp.Left = 53
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                    txtScan.Text = ""
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    'InitializeForm(4)
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    If MessageBox.Show("Are You Sure To Delete Trip No" + txtTripNo.Text.Trim() + " ?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        If DeleteData() = True Then
                            InitializeForm(1)
                            cmdGrp.Revert()
                            cmdGrp.Top = 14
                            cmdGrp.Left = 53
                            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                        End If
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        Finally

        End Try
    End Sub
    Private Function DeleteData() As Boolean
        Dim strSql As String = String.Empty
        Dim flag As Boolean = True
        Try
            If isTripUsed(txtTripNo.Text) Then
                MessageBox.Show("Trip cannot be deleted .It is used in Freight Bill .", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                flag = False
                Exit Function
            End If
            strSql = "Delete from Freight_Trip_STAFF_Hdr  WHERE Unit_Code='" & gstrUNITID & "' AND Trip_No='" & txtTripNo.Text & "'"
            SqlConnectionclass.ExecuteNonQuery(strSql)
            Return True
        Catch ex As Exception
            flag = False
            RaiseException(ex)
        End Try
        Return flag
    End Function

    Private Function SaveData() As Boolean
        Try
            Dim strSql As String = String.Empty
            Dim contractType As String = String.Empty

            If String.IsNullOrEmpty(txtRoute.Text.Trim()) Then
                MessageBox.Show("Route No cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtRoute.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(txtShift.Text.Trim()) Then
                MessageBox.Show("Shift No cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtShift.Focus()
                Return False
            End If

            If Not isContractExist(txtRoute.Text, txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                MessageBox.Show("Contract not defined for selected Route.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Return False
            End If

            Dim HoursAllowedNextStaffTrip As Integer = 0
            HoursAllowedNextStaffTrip = SqlConnectionclass.ExecuteScalar("SELECT HoursAllowedNextStaffTrip from POConfig_Mst where UNIT_CODE='" + gstrUNITID + "'")

            Dim strDate As Date?
            strDate = SqlConnectionclass.ExecuteScalar("Select  top 1 dateadd(HOUR," & HoursAllowedNextStaffTrip & ",Trip_Date) from Freight_Trip_STAFF_Hdr  where Docno='" & txtDocNo.Text & "' and route_code='" & txtRoute.Text & "'  and transporter_code='" & txtTransporterCode.Text & "' and Vehicle_Category_Code='" & txtVehicleCategory.Text & "' and Vehicle_No='" & txtVehicleNo.Text & "' order by Ent_Dt desc ")
            If GetServerDateTime() < strDate Then
                MessageBox.Show("Trip is already created for these details.You cannot create trip before " & strDate, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Return False
            End If





            If MessageBox.Show("Are you sure, you want to save?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                Return False
            End If


            Dim cmd As SqlCommand = New SqlCommand
            Dim strDocstring As String = ""
            With cmd
                .Connection = SqlConnectionclass.GetConnection()
                .Transaction = .Connection.BeginTransaction
                .CommandType = CommandType.StoredProcedure

                If Len(txtTripNo.Text.Trim) = 0 Then
                    dtpTripDt.Value = GetServerDateTime()
                End If

                .CommandText = "USP_FREIGHT_TRIP_STAFF_GENERATION"
                .Parameters.Clear()
                .Parameters.AddWithValue("@p_UnitCode", gstrUNITID)
                .Parameters.Add(New SqlParameter("@P_TripNo", SqlDbType.VarChar, 500, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                .Parameters("@P_TripNo").Value = txtTripNo.Text.Trim()
                .Parameters.AddWithValue("@p_TripDt", dtpTripDt.Value.ToString("dd MMM yyyy HH:mm"))
                .Parameters.AddWithValue("@P_TransporterCode", txtTransporterCode.Text.Trim())
                .Parameters.AddWithValue("@P_VehicleCategory", txtVehicleCategory.Text.Trim())
                .Parameters.AddWithValue("@P_VehicleNo", txtVehicleNo.Text.Trim())
                .Parameters.AddWithValue("@p_TranType", "SAVE")
                .Parameters.AddWithValue("@p_UserId", mP_User)
                .Parameters.AddWithValue("@ROUTE_CODE", txtRoute.Text)
                .Parameters.AddWithValue("@SHIFT_CODE", txtShift.Text)
                .Parameters.AddWithValue("@DOCNO", txtDocNo.Text)
                .Parameters.AddWithValue("@SEATING_CAPACITY", txtSeatingCapacity.Text)
                .Parameters.Add(New SqlParameter("@P_Error", SqlDbType.VarChar, 500, ParameterDirection.Output, True, 0, 0, "", DataRowVersion.Default, ""))

                .ExecuteNonQuery()
                If String.IsNullOrEmpty(.Parameters("@P_Error").Value) Then
                    txtTripNo.Text = .Parameters("@P_TripNo").Value.ToString()
                    .Transaction.Commit()

                    MessageBox.Show("Data saved successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return True
                Else
                    .Transaction.Rollback()
                    MessageBox.Show(.Parameters("@P_Error").Value, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
            End With

        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return False
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

    Private Function isContractExist(ByVal strRouteCode As String, ByVal strTransporterCode As String, ByVal effDate As String) As Boolean
        Try
            Return IsRecordExists("SELECT * FROM FREIGHT_CONTRACT_STAFF FCD " & _
                            " INNER JOIN FREIGHT_CONTRACT_HDR FCH ON FCD.CONTRACT_ID=FCH.CONTRACT_ID AND FCD.UNIT_CODE=FCH.UNIT_CODE" & _
                            " AND FCH.TRANSPORTER_CODE='" + strTransporterCode + "' AND FCH.AUTHORIZED_STATUS='AUTHORIZED' AND FCH.CONTRACT_TYPE='" + _contractTypeStaffBased + "' " & _
                            " WHERE FCD.Route_Code='" + strRouteCode + "' AND FCD.UNIT_CODE='" + gstrUNITID + "' AND '" + effDate + "' BETWEEN FCD.EFFECTIVE_FROM AND FCD.EFFECTIVE_TO " & _
                            " AND FCD.VEHICLE_CATEGORY_CODE='" + txtVehicleCategory.Text.Trim() + "' AND FCH.CONTRACT_TYPE='Staff' AND ISNULL(FCD.CONTRACT_STATUS,'')<>'AUTHORIZED' ")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Function isTripUsed(ByVal strTripNo As String) As Boolean
        Try
            Return IsRecordExists(" SELECT * FROM FREIGHT_STAFFBUS_SUBMISSION_dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' and Trip_No='" & strTripNo & "'")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
    Private Sub BtnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpTrip.Click, BtnHelpShift.Click, btnHelpRoute.Click
        Dim strSQL As String = String.Empty
        Dim strHelp As String()
        Try
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            If sender Is BtnHelpTrip Then
                strSQL = "SELECT Trip_No, convert(varchar(20),Trip_Date,113) TripDate, Transporter_Code, Vehicle_Category_Code, Vehicle_No, AuthStatus " & _
                        "  from Freight_Trip_STAFF_Hdr where Unit_Code='" + gstrUNITID + "' ORDER BY Trip_Date DESC "

                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Trip", 1, 0, txtTripNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length >= 5 Then
                        txtTripNo.Text = strHelp(0).Trim
                        dtpTripDt.Value = strHelp(1)
                        strSQL = "Select TRIPHDR.DocNo,Trip_No, convert(varchar(20),Trip_Date,106) TripDate, ISNULL(Transporter_Code,'') Transporter_Code, " & _
                        "ISNULL(Vehicle_Category_Code,'') Vehicle_Category_Code, Vehicle_No,SEATING_CAPACITY ,TRIPHDR.ROUTE_CODE, Route_Desc,Place_From,Place_To,  " & _
                        "Distance_KM,MAX_THRESHOLD_KM,TRIPHDR.Shift_code, convert(char(5),CAST(start_tm as datetime),108) as start_tm,  " & _
                        "convert(char(5),CAST(end_tm as datetime),108) as end_tm   from Freight_Trip_STAFF_Hdr TRIPHDR  " & _
                        "inner join VEHICLE_ROUTE_MASTER VHROUTE on TRIPHDR.Unit_Code=VHROUTE.UNIT_CODE and TRIPHDR.Docno=VHROUTE.Docno  " & _
                        "and TRIPHDR.ROUTE_CODE=VHROUTE.ROUTE_CODE Inner join shift_mst SHIFT on TRIPHDR.Unit_Code=Shift.unit_code  " & _
                        "and TRIPHDR.SHIFT_CODE=Shift.SHift_code WHERE TRIPHDR.Unit_Code='" + gstrUNITID + "' AND Trip_No='" + strHelp(0).Trim() + "'"
                        Dim dt As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                        If (dt.Rows.Count > 0) Then
                            txtDocNo.Text = Convert.ToString(dt.Rows(0)("Transporter_Code"))
                            txtTransporterCode.Text = Convert.ToString(dt.Rows(0)("Transporter_Code"))
                            txtTransporterName.Text = SqlConnectionclass.ExecuteScalar(" SELECT VM.VENDOR_NAME FROM VENDOR_MST VM WHERE VM.UNIT_CODE='" + gstrUNITID + "' AND VM.VENDOR_CODE = '" + txtTransporterCode.Text + "'")
                            txtVehicleCategory.Text = Convert.ToString(dt.Rows(0)("Vehicle_Category_Code"))
                            txtVehicleTypedesc.Text = SqlConnectionclass.ExecuteScalar("Select Description from FREIGHT_VEHICLE_VENDOR_CATEGORY_MST where vehicle_category_code='" & txtVehicleCategory.Text & "' and unit_code='" & gstrUNITID & "'")
                            txtVehicleNo.Text = Convert.ToString(dt.Rows(0)("Vehicle_No"))
                            txtSeatingCapacity.Text = Convert.ToString(dt.Rows(0)("SEATING_CAPACITY"))
                            txtRoute.Text = Convert.ToString(dt.Rows(0)("ROUTE_CODE"))
                            txtRouteDesc.Text = Convert.ToString(dt.Rows(0)("Route_Desc"))
                            txtPlaceFrom.Text = Convert.ToString(dt.Rows(0)("Place_From"))
                            txtPlaceTo.Text = Convert.ToString(dt.Rows(0)("Place_To"))
                            txtKm.Text = Convert.ToString(dt.Rows(0)("MAX_THRESHOLD_KM"))
                            txtThresholdKm.Text = Convert.ToString(dt.Rows(0)("Transporter_Code"))
                            txtShift.Text = Convert.ToString(dt.Rows(0)("Shift_code"))
                            txtStartTime.Text = Convert.ToString(dt.Rows(0)("start_tm"))
                            txtEndTime.Text = Convert.ToString(dt.Rows(0)("end_tm"))
                        End If
                       


                        InitializeForm(3)
                    End If
                End If
            ElseIf sender Is BtnHelpShift Then
                strSQL = " Select Shift_code,  convert(char(5),CAST(start_tm as datetime),108) as start_tm, convert(char(5),CAST(end_tm as datetime),108) as end_tm from shift_mst  where unit_code='" & gstrUNITID & "'"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Shift", 1, 0, txtTripNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 AndAlso Not IsNothing(strHelp(1)) Then
                        txtShift.Text = strHelp(0).Trim
                        txtStartTime.Text = strHelp(1).Trim
                        txtEndTime.Text = strHelp(2).Trim
                    End If
                End If
            ElseIf sender Is btnHelpRoute Then
                strSQL = " select Route_Code,Route_Desc,Place_From,Place_To,Distance_KM,MAX_THRESHOLD_KM  from VEHICLE_ROUTE_MASTER  where  DocNo='" & txtDocNo.Text & "' and unit_code='" & gstrUNITID & "'"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Route", 1, 0, txtTripNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 AndAlso Not IsNothing(strHelp(1)) Then
                        txtRoute.Text = strHelp(0).Trim
                        txtRouteDesc.Text = strHelp(1).Trim
                        txtPlaceFrom.Text = strHelp(2).Trim
                        txtPlaceTo.Text = strHelp(3).Trim
                        txtKm.Text = strHelp(4).Trim
                        txtThresholdKm.Text = strHelp(5).Trim
                    End If
                End If
            End If


        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtSplChars_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTripNo.KeyPress
        Try
            If e.KeyChar = "'" Then
                e.Handled = True
            End If
            If sender Is txtVehicleNo Then
                If Not ((e.KeyChar >= "A" And e.KeyChar <= "Z") Or (e.KeyChar >= "a" And e.KeyChar <= "z") Or (e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = "-" Or System.Text.Encoding.ASCII.GetBytes(e.KeyChar)(0) = 8) Then
                    e.Handled = True
                End If
            
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtTripNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            If Not IsNumeric(e.KeyChar) Then
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtTripNo_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTripNo.Validating

    End Sub

    Private Sub txt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTripNo.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                If sender Is txtTripNo And cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    BtnHelp_Click(BtnHelpTrip, New EventArgs())
                ElseIf cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If sender Is txtTransporterCode Then

                    ElseIf sender Is txtVehicleCategory Then

                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub



    Private Sub fillDocNo()
        Try
            Dim strSQL, strSelectedRec As String
            Dim strHelp As String()
            Dim rNo As Integer = 1
            Dim noOfDays As Integer = 0
            Dim errMsg As String = String.Empty
            Dim strEwayBillType As String = ""
            strSQL = String.Empty
            'CheckCodeExistence Contract query from trip screen
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub InitializeForm(ByVal Form_Status_flag As Integer)
        Try
            If Form_Status_flag = 1 Then   'Page Load
                txtTripNo.Text = ""
                txtScan.Text = ""
                txtTripNo.Enabled = True
                txtTripNo.ReadOnly = False
                dtpTripDt.Enabled = False
                dtpTripDt.Value = GetServerDate()
                txtTransporterCode.ReadOnly = True
                txtTransporterCode.Text = ""
                txtTransporterName.Text = String.Empty
                txtVehicleCategory.ReadOnly = True
                txtVehicleCategory.Text = ""
                txtVehicleNo.Text = String.Empty
                txtVehicleNo.ReadOnly = True
                txtVehicleTypedesc.Text = ""
                txtVehicleTypedesc.ReadOnly = True
                txtSeatingCapacity.Text = ""
                txtShift.Text = ""
                BtnHelpTrip.Enabled = True
                BtnHelpShift.Enabled = True
                btnHelpRoute.Enabled = True
                txtRoute.Text = ""
                txtRouteDesc.Text = ""
                txtPlaceFrom.Text = ""
                txtPlaceTo.Text = ""
                txtKm.Text = ""
                txtThresholdKm.Text = ""
                txtStartTime.Text = ""
                txtEndTime.Text = ""
                BtnHelpShift.Enabled = False
                btnHelpRoute.Enabled = False
                txtRoute.ReadOnly = True
                txtRouteDesc.ReadOnly = True
                txtPlaceFrom.ReadOnly = True
                txtPlaceTo.ReadOnly = True
                txtKm.ReadOnly = True
                txtThresholdKm.ReadOnly = True
                txtStartTime.ReadOnly = True
                txtEndTime.ReadOnly = True
               
                cmdGrp.Revert()
                cmdGrp.Top = 14
                cmdGrp.Left = 53
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
            ElseIf Form_Status_flag = 2 Then    'Add Mode
                BtnHelpTrip.Enabled = False
                txtTripNo.Enabled = False
                txtTripNo.Text = String.Empty
                dtpTripDt.Enabled = False
                txtVehicleNo.Text = String.Empty
                txtVehicleNo.ReadOnly = Not (cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD)
                txtTransporterCode.Text = String.Empty
                txtTransporterName.Text = String.Empty
                txtVehicleCategory.Text = String.Empty             
                BtnHelpShift.Enabled = True
                btnHelpRoute.Enabled = True
                txtRoute.Text = ""
                txtRouteDesc.Text = ""
                txtPlaceFrom.Text = ""
                txtPlaceTo.Text = ""
                txtKm.Text = ""
                txtThresholdKm.Text = ""
                txtStartTime.Text = ""
                txtEndTime.Text = ""
                txtVehicleTypedesc.Text = ""
                txtSeatingCapacity.Text = ""
                txtShift.Text = ""
                dtpTripDt.Value = GetServerDateTime()
                dtpTripDt.Focus()
                txtScan.Focus()
            ElseIf Form_Status_flag = 3 Then    'View Mode
                dtpTripDt.Enabled = False
                BtnHelpTrip.Enabled = True
              
                'If IsRecordExists(" SELECT Trip_No FROM Freight_Gate_Outward_Reg_Hdr WHERE Unit_code='" + gstrUNITID + "' and Trip_No='" + txtTripNo.Text.Trim() + "' AND CANCEL_YN = 'N'") Then
                '    '101053952-starts
                '    'If IsRecordExists(" SELECT Trip_No FROM Freight_Gate_Outward_Reg_Hdr WHERE Unit_code='" + gstrUNITID + "' and Trip_No='" + txtTripNo.Text.Trim() + "' AND CANCEL_YN = 'N'" & _
                '    '                       " UNION " & _
                '    '                       " SELECT Trip_No FROM Freight_Trip_Gen_Hdr " & _
                '    '                       " WHERE Unit_Code='" + gstrUNITID + "' AND Trip_No='" + txtTripNo.Text.Trim() + "' and IsTransporterVehicle=0 and ISNULL(AuthStatus,'')='A'") Then
                '    '101053952-ends
                '    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                '    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                'Else
                '    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                '    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                'End If
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
            ElseIf Form_Status_flag = 4 Then
                BtnHelpTrip.Enabled = False
                txtTripNo.Enabled = True
                txtTripNo.ReadOnly = True
                txtVehicleNo.ReadOnly = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtScan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtScan.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Dim strQuery As String = String.Empty
        Dim strDocNo As String = String.Empty
        Dim strTripNo As String = String.Empty
        Dim oSqlDr As SqlDataReader = Nothing

        Try
            If (KeyAscii = Keys.Enter) AndAlso txtTripNo.Text = "" Then
                Dim arr As String()
                Dim strScanTrip As String
                Dim strTransporter As String
                arr = txtScan.Text.Trim.Split("|")
                If arr.Count > 1 Then
                    strScanTrip = arr(0).ToString()
                    strTransporter = arr(1).ToString()
                End If


                strQuery = " SELECT * " & _
                           " FROM vehicle_mst  WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOCNO = '" & strScanTrip & "'  and transporter_code='" & strTransporter & "' "
                oSqlDr = SqlConnectionclass.ExecuteReader(strQuery)

                If (oSqlDr.HasRows = True) Then
                    oSqlDr.Read()
                    txtDocNo.Text = strScanTrip.Trim
                    txtTransporterCode.Text = oSqlDr("TRANSPORTER_CODE").ToString.Trim
                    txtTransporterName.Text = oSqlDr("TRANSPORTER_NAME").ToString.Trim
                    txtVehicleCategory.Text = oSqlDr("VEHICLE_TYPE").ToString.Trim
                    txtVehicleTypedesc.Text = SqlConnectionclass.ExecuteScalar("Select Description from FREIGHT_VEHICLE_VENDOR_CATEGORY_MST where vehicle_category_code='" & txtVehicleCategory.Text & "' and unit_code='" & gstrUNITID & "'")
                    txtVehicleNo.Text = oSqlDr("VEHICLE_NO").ToString.Trim
                    txtSeatingCapacity.Text = oSqlDr("SEATING_CAPACITY").ToString.Trim
                Else
                    MessageBox.Show("Data does not exist for barcode", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
                If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
                txtScan.Text = ""

            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally

        End Try
    End Sub


End Class