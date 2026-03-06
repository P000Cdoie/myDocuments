Imports System.Data
Imports System.Data.SqlClient

'=======================================================================================
'Copyright          :   Mothersonsumi Infotech & Design Ltd.
'Module             :   Stores
'Author             :   Priti Sharma
'Creation Date      :   23 Oct 2019
'Description        :   Exception Approval 
'=======================================================================================


Public Class frmMKTTRN0124

    Enum GridEnum
        Status = 1
        TripNo
        TripDate
        Route
        VehicleCat
        Distance
        ExtraKm
    End Enum

   

    Dim DocFrm As Form
    Dim dtDocTable As DataTable


    Private Sub AddColumns()
        With GridDtl
            .MaxRows = 0
            .MaxCols = [Enum].GetNames(GetType(GridEnum)).Count
            .Row = 0
            .set_RowHeight(0, 25)
            .Col = 0 : .set_ColWidth(0, 3)
            .Col = GridEnum.Status : .Text = "Status" : .set_ColWidth(GridEnum.Status, 4)
            .Col = GridEnum.TripNo : .Text = "Trip No" : .set_ColWidth(GridEnum.TripNo, 30)
            .Col = GridEnum.TripDate : .Text = "Trip Date" : .set_ColWidth(GridEnum.TripDate, 30)
            .Col = GridEnum.Route : .Text = "Route" : .set_ColWidth(GridEnum.Route, 12)
            .Col = GridEnum.VehicleCat : .Text = "Vehicle Cat" : .set_ColWidth(GridEnum.VehicleCat, 12)
            .Col = GridEnum.Distance : .Text = "Distance" : .set_ColWidth(GridEnum.Distance, 12)
            .Col = GridEnum.ExtraKm : .Text = "Extra Km" : .set_ColWidth(GridEnum.ExtraKm, 12)

        End With

    End Sub

 
    Private Sub AddBlankRow()
        Try
            With Me.GridDtl
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 15)
                .Col = GridEnum.Status : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .Col = GridEnum.TripNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = GridEnum.TripDate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = GridEnum.Route : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = GridEnum.VehicleCat : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = GridEnum.Distance : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = GridEnum.ExtraKm : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
  

    Private Sub Clear()
        txtTransportCode.Text = ""
        txtTransporterName.Text = ""
        dtpFromDate.Value = GetServerDate()
        dtpToDate.Value = GetServerDate()
        GridDtl.MaxRows = 0

    End Sub




    Private Sub FRMSTRTRN0044_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Try
            If e.KeyChar = "'" Then
                e.Handled = True
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub FRMSTRTRN0044_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            
            Call FitToClient(Me, GrpMain, ctlFormHeader, GrpBoxSave)
            GrpMain.Height = GrpMain.Height + 30
            Me.MdiParent = mdifrmMain
            'GrpMain.Height = GrpMain.Height + 30
            cmdButtons.Top = 560
            cmdButtons.Left = 400

            GridDtl.MaxRows = 0
            EnableDisable()
            AddColumns()
            'Me.KeyPreview = True



        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub


    Private Sub EnableDisable()

        Try
            'If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            '    txtDocNo.Enabled = False
            '    cmdDocNo.Enabled = False
            '    TxtRemarks.Enabled = True
            '    cmdTrip_No.Enabled = True
            'Else
            '    cmdTrip_No.Enabled = False

            'End If
        Catch Ex As Exception
            RaiseException(Ex)
        End Try

    End Sub



    Private Sub TxtRemarks_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Try
            If Asc(e.KeyChar()) = 34 Or Asc(e.KeyChar()) = 39 Then
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub



    
    Private Sub GetGrid_Data()
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim DataRd As SqlDataReader
        Dim strSql As String = String.Empty
        Try
            strSql = "Select Trip_no,Trip_Date,Transporter_Code,Vehicle_Category_Code,TRIPHDR.ROUTE_CODE,DISTANCE_KM,ExtraKM from Freight_Trip_STAFF_Hdr TRIPHDR  " & _
            "inner join VEHICLE_ROUTE_MASTER VHROUTE on TRIPHDR.Unit_Code=VHROUTE.UNIT_CODE and TRIPHDR.Docno=VHROUTE.Docno and " & _
            "TRIPHDR.ROUTE_CODE=VHROUTE.ROUTE_CODE where  isnull(AuthStatus,0)=0 and TRIPHDR.unit_code='" & gstrUNITID & "' and  convert(date, Trip_date,103)  between '" & Format(dtpFromDate.Value, "dd MMM yyyy") & "' and '" & Format(dtpToDate.Value, "dd MMM yyyy") & "'  and transporter_code='" & txtTransportCode.Text & "'" & _
            "and trip_no not in (select trip_no from FREIGHT_STAFFBUS_SUBMISSION_dtl   where unit_code='" & gstrUNITID & "')"
            Dim dt As DataTable = SqlConnectionclass.GetDataTable(strSql)
            If (dt.Rows.Count > 0) Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    With GridDtl
                        AddBlankRow()
                        .Col = 2 : .Col2 = 5
                        .Row = .MaxRows : .Row2 = .MaxRows
                        .BlockMode = True : .BlockMode = False
                        .Col = 2 : .Col2 = 5
                        .Row = .MaxRows : .Row2 = .MaxRows
                        .BlockMode = True : .Lock = True : .BlockMode = False
                        .SetText(GridEnum.TripNo, i + 1, dt.Rows(i).Item("Trip_no").ToString.Trim)
                        .SetText(GridEnum.TripDate, i + 1, dt.Rows(i).Item("Trip_Date").ToString.Trim)
                        .SetText(GridEnum.VehicleCat, i + 1, dt.Rows(i).Item("Vehicle_Category_Code").ToString.Trim)
                        .SetText(GridEnum.Route, i + 1, dt.Rows(i).Item("ROUTE_CODE").ToString.Trim)
                        .SetText(GridEnum.Distance, i + 1, dt.Rows(i).Item("DISTANCE_KM").ToString.Trim)
                        .SetText(GridEnum.ExtraKm, i + 1, dt.Rows(i).Item("ExtraKM").ToString.Trim)
                    End With
                Next
            Else
                MessageBox.Show("No Data avaiable.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

            End If

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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

   
 

    Private Function Save_Data(ByVal strType As String) As Boolean
        Dim Sqlcmd As New SqlCommand
        Dim intY As Integer
        Dim intX As Integer
        Dim strSql As String = String.Empty
        Dim intExtraKm As Integer = 0
        Dim SqlTrans As SqlTransaction
        Dim strTripNo As String = String.Empty
        Dim blnStatus As Boolean = False
        Try
            With GridDtl
                For intRow As Integer = 1 To .MaxRows
                    .Row = intRow
                    .Col = GridEnum.Status
                    blnStatus = Val(.Value)
                    .Col = GridEnum.TripNo
                    strTripNo = (.Text)
                    .Col = GridEnum.ExtraKm
                    intExtraKm = Val(.Value)
                    If strType = "A" Then
                        If blnStatus Then
                            strSql = "Update Freight_Trip_STAFF_Hdr set AuthStatus=1,ExtraKm=" & intExtraKm & " where Trip_no='" & strTripNo & "' and unit_code='" & gstrUNITID & "'"
                            SqlConnectionclass.ExecuteNonQuery(strSql)
                          
                        End If
                    ElseIf strType = "R" Then
                        If blnStatus Then
                            strSql = "Update Freight_Trip_STAFF_Hdr set RejStatus=1 where Trip_no='" & strTripNo & "' and unit_code='" & gstrUNITID & "'"
                            SqlConnectionclass.ExecuteNonQuery(strSql)

                        End If
                    End If
                  
                Next
                If strType = "A" Then
                    MessageBox.Show("Data Authorised successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Else
                    MessageBox.Show("Data Rejected successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                End If
            End With
        Catch ex As Exception
            If Not IsNothing(SqlTrans) Then
                SqlTrans.Rollback()
            End If
            RaiseException(ex)
        End Try
    End Function
  



       Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    
    Private Sub btnApprove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        Save_Data("A")
        Clear()
    End Sub

    Private Sub btnShowData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowData.Click
        If txtTransportCode.Text = "" Then
            MessageBox.Show("Select Transporter first.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Else
            GridDtl.MaxRows = 0
            GetGrid_Data()
        End If
       
    End Sub

    Private Sub btnHelpTransport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnHelpTransport.Click
        Dim strSQL As String = String.Empty
        Dim strHelp As String()
        Try
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            strSQL = " SELECT VM.VENDOR_CODE, VM.VENDOR_NAME FROM VENDOR_MST VM WHERE VM.TRANSPORTER_FLAG=1 AND VM.ACTIVE_FLAG='A' AND VM.UNIT_CODE='" + gstrUNITID + "' ORDER BY VENDOR_CODE"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Transporter", 1, 0, "")
            If Not IsNothing(strHelp) Then
                If strHelp.Length > 0 AndAlso Not IsNothing(strHelp(1)) Then
                    txtTransportCode.Text = strHelp(0).Trim
                    txtTransporterName.Text = strHelp(1).Trim()                   
                    clearGrid()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub clearGrid()
        Try
            GridDtl.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnReject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReject.Click
        Save_Data("R")
        Clear()
    End Sub

    Private Sub btnReturn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReturn.Click

    End Sub
End Class