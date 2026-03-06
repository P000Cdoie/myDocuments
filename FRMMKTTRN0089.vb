Imports System
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Collections.Generic
Imports VB = Microsoft.VisualBasic
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices


Public Class FRMMKTTRN0089
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
    'Modified By        - Praveen Kumar
    'Created Date       - 31 July 2024
    'description        - Save QR Code image 
    '=========================================================================================================

#Region "Global variables"
    Dim dtDocDtl As DataTable
    Dim REPDOC As ReportDocument
    Dim REPVIEWER As eMProCrystalReportViewer
    Const _contractTypeTripBased As String = "TRIP"
    Const _contractTypeMonthlyBased As String = "MONTHLY"
    Const _contractTypeCourierBased As String = "SIZE_WEIGHT"
    Dim mblnAllowTransporterfromMaster As Boolean = False
    Dim blnIsAdhoc As Boolean
    Dim CustomerCode As String = String.Empty
    Dim CustomerCodes As New List(Of String)()


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
    Private Sub FRMMKTTRN0089_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If IsRecordExists("select IsFreightActive from POConfig_Mst where UNIT_CODE='" + gstrUNITID + "' and IsFreightActive=0") Then
                MessageBox.Show("This functionality is not relevant to you.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Me.Close()
                Return
            End If
            Call FitToClient(Me, GrpMain, ctlFormHeader, GrpCmdBtn, 500)
            Me.MdiParent = mdifrmMain
            InitializeForm(1)
            'added by priti on 16 march 2020 to add vehicle help box
            mblnAllowTransporterfromMaster = CBool(Find_Value("SELECT isnull(AllowTransporterfromMaster,0) as AllowTransporterfromMaster  FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
            If mblnAllowTransporterfromMaster Then
                'txtVehicleNo.Enabled = True
                cmdVehicleCodeHelp.Visible = True
                txtVehicleCategory.Enabled = False
                BtnHelpVehCategory.Enabled = False
            Else
                txtVehicleNo.Enabled = True
                cmdVehicleCodeHelp.Visible = False
                txtVehicleCategory.Enabled = True
                BtnHelpVehCategory.Enabled = True

            End If

            Me.BringToFront()

            setAdHoc()

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region

#Region "Control Events"

    Private Function AdhocTripGenerationAutomail() As Boolean
        Dim Row As DataRow
        Dim intCell As Integer
        Dim intRow As Integer
        Dim blnRecordExists As Boolean = False
        Try
            SqlConnectionclass.CloseGlobalConnection()
            SqlConnectionclass.OpenGlobalConnection()
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            SqlConnectionclass.BeginTrans()
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_AUTOMAIL_Adhoc_Trip_Generation"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@Request_No", SqlDbType.Int).Value = txtReqNo.Text
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 16).Value = mP_User
                    .Parameters.Add("@Trip_No", SqlDbType.VarChar, 16).Value = txtTripNo.Text
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                    sqlcmd.Dispose()
                End With
            End Using
            SqlConnectionclass.CommitTran()

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)


            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Function

    Private Sub cmdGrp_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdGrp.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    InitializeForm(2)
                    'added by priti on 16 march 2020 to add vehicle help box
                    If mblnAllowTransporterfromMaster Then
                        'txtVehicleNo.Enabled = True
                        cmdVehicleCodeHelp.Visible = True
                        cmdVehicleCodeHelp.Enabled = True
                        txtVehicleCategory.Enabled = False
                        BtnHelpVehCategory.Enabled = False
                    Else
                        txtVehicleNo.Enabled = True
                        cmdVehicleCodeHelp.Visible = False
                        cmdVehicleCodeHelp.Enabled = False
                        txtVehicleCategory.Enabled = True
                        BtnHelpVehCategory.Enabled = True
                    End If
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If SaveData() Then
                        If rdoAdHoc.Checked = True Then
                            AdhocTripGenerationAutomail()
                        End If
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
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    InitializeForm(4)
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
                    If String.IsNullOrEmpty(txtTripNo.Text.Trim()) Then
                        MessageBox.Show("Select Trip No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        txtTripNo.Focus()
                        Return
                    End If

                    ''Praveen Start QRCODE GENERATION AND SAVE. 31.07.2024
                    Dim strBarCode As String
                    Dim strBarcodeMsg As String = ""
                    Dim strLableBarcode As String = ""
                    Dim strCustBarcode As String = ""
                    strBarCode = SqlConnectionclass.ExecuteScalar("SELECT TRIP_BARCODE FROM FREIGHT_TRIP_GEN_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND TRIP_NO='" + txtTripNo.Text.Trim() + "' ")
                    Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUNITID)
                    strBarcodeMsg = ObjBarcodeHMI.GenerateVehicleBarCode(gstrUserMyDocPath, strBarCode, gstrCONNECTIONSTRING)
                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                        MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                        Exit Sub
                    End If

                    If SaveBarCodeImage(strBarcodeMsg, gstrUserMyDocPath) = False Then
                        MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                        Exit Sub
                    End If
                    ''PRAVEEN END QRCODE GENERATION AND SAVE.

                    REPVIEWER = New eMProCrystalReportViewer()
                    REPDOC = REPVIEWER.GetReportDocument
                    Dim strReportName As String = String.Empty
                    If ALLOWCUSTOMERSPECIFICREPORT_ROWSIZELARGE() Then
                        strReportName = "\Reports\rptTripGeneration_Rowsizelarge.rpt"
                    Else
                        strReportName = "\Reports\rptTripGeneration.rpt"
                    End If

                    Dim strRepPath As String = My.Application.Info.DirectoryPath & strReportName
                    REPDOC.Load(strRepPath)
                    REPDOC.RecordSelectionFormula = " {Vw_TripDetail.Unit_Code}='" + gstrUNITID + "' and {Vw_TripDetail.Trip_No}='" + txtTripNo.Text.Trim() + "'"
                    REPDOC.DataDefinition.FormulaFields("CompanyName").Text = "'" + gstrCOMPANY + "'"
                    REPDOC.DataDefinition.FormulaFields("CompanyAddress").Text = "'" + gstr_WRK_ADDRESS1 + "'"
                    REPDOC.DataDefinition.FormulaFields("CompanyAdd2").Text = "'" + gstr_WRK_ADDRESS2 + "'"
                    REPVIEWER.Show()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        Finally

        End Try
    End Sub

    Public Function SaveBarCodeImage(ByVal VehicleBarCode As String, ByVal pstrPath As String) As Boolean
      
        On Error GoTo ErrHandler
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        Dim strBarcodeMsg_paratemeter As String
        Dim objConnection As New prj_Connection.cls_Connection
        Dim blnCROP_QRIMAGE As Boolean = False

        SaveBarCodeImage = True
        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        stimage.Open()
        pstrPath = pstrPath & "\BarcodeImg.wmf"
        blnCROP_QRIMAGE = CBool(Find_Value("SELECT CROP_QRBARCODE  FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        If blnCROP_QRIMAGE = True Then
            Dim bmp As New Bitmap(pstrPath)
            Dim picturebox1 As New PictureBox
            picturebox1.Image = ImageTrim(bmp)
            picturebox1.Image.Save(pstrPath)
            picturebox1 = Nothing
        End If
        stimage.LoadFromFile(pstrPath)

        strQuery = "select  Trip_Barcode_QR_Image from Freight_Trip_Gen_Hdr where Trip_No='" & txtTripNo.Text & "'  and Unit_Code = '" & gstrUNITID & "'"
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Rs.Fields("Trip_Barcode_QR_Image").Value = stimage.Read

        Rs.Update()
        Rs.Close()
        Rs = Nothing

        Exit Function
ErrHandler:
        SaveBarCodeImage = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
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

    Private Sub BtnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpTrip.Click, BtnHelpTransporter.Click, BtnHelpVehCategory.Click
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
                strSQL = "SELECT Trip_No, convert(varchar(20),Trip_Date,113) TripDate, Transporter_Code, Vehicle_Category_Code, Vehicle_No " & _
                        " , CASE WHEN IsTransporterVehicle=1 THEN 'Transporter' ELSE 'Customer Own' END [Transporter/Customer Own], AuthStatus,adhocReqNo, " & _
                        " Trip_Date from Freight_Trip_Gen_Hdr where Unit_Code='" + gstrUNITID + "'" & _
                        " Union " & _
                        "SELECT Trip_No, convert(varchar(20),Trip_Date,113) TripDate, Transporter_Code,Vehicle_Category Vehicle_Category_Code, Lorry_No Vehicle_No " & _
                        " , 'Transporter' [Transporter/Customer Own], 'A' AuthStatus, RequestNo, " & _
                        " Trip_Date from Freight_Gate_Inward where Unit_Code='" + gstrUNITID + "' ORDER BY Trip_Date DESC "

                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Trip", 1, 0, txtTripNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length >= 5 Then
                        txtTripNo.Text = strHelp(0).Trim
                        dtpTripDt.Value = strHelp(1)
                        strSQL = "SELECT Trip_No, convert(varchar(20),Trip_Date,106) TripDate, ISNULL(Transporter_Code,'') Transporter_Code, ISNULL(Vehicle_Category_Code,'') Vehicle_Category_Code, Vehicle_No " & _
                        " , IsTransporterVehicle, isnull(CarrierName,'') CarrierName, isnull(VehicleType,'') VehicleType, isnull(AuthBy,'') AuthBy, AuthStatus" & _
                        ", ISNULL(CONTRACT_TYPE,'') CONTRACT_TYPE, ISNULL(SIZE_LENGTH_FT,0) SIZE_LENGTH_FT, ISNULL(SIZE_LENGTH_INCH,0) SIZE_LENGTH_INCH" & _
                        ", ISNULL(SIZE_WIDTH_FT,0) SIZE_WIDTH_FT, ISNULL(SIZE_WIDTH_INCH,0) SIZE_WIDTH_INCH, ISNULL(SIZE_HEIGHT_FT,0) SIZE_HEIGHT_FT" & _
                        ", ISNULL(SIZE_HEIGHT_INCH,0) SIZE_HEIGHT_INCH, ISNULL(SIZE_CUBIC_FT,0) SIZE_CUBIC_FT, ISNULL(WEIGHT_KG,0) WEIGHT_KG, ISNULL(ZONE,'') ZONE,ISNULL(TRANSPORTER_MODE,'') TRANSPORTER_MODE, adhocReqNo " & _
                        " FROM Freight_Trip_Gen_Hdr WHERE Unit_Code='" + gstrUNITID + "' AND Trip_No='" + strHelp(0).Trim() + "'" & _
                        " Union" & _
                        " SELECT Trip_No, convert(varchar(20),Trip_Date,106) TripDate, ISNULL(Transporter_Code,'') Transporter_Code, '' Vehicle_Category_Code, lorry_no Vehicle_No" & _
                        " , '' IsTransporterVehicle, '' CarrierName, '' VehicleType, '' AuthBy, 'A' AuthStatus" & _
                        " , '' CONTRACT_TYPE, 0 SIZE_LENGTH_FT, 0 SIZE_LENGTH_INCH, 0 SIZE_WIDTH_FT, 0 SIZE_WIDTH_INCH, 0 SIZE_HEIGHT_FT" & _
                        " , 0 SIZE_HEIGHT_INCH, 0 SIZE_CUBIC_FT, 0 WEIGHT_KG, '' ZONE,'' TRANSPORTER_MODE, RequestNo adhocReqNo" & _
                        " FROM Freight_Gate_Inward WHERE Unit_Code='" + gstrUNITID + "' AND Trip_No = '" + strHelp(0).Trim() + "'"

                        Dim dt As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                        If (dt.Rows.Count > 0) Then
                            If (Convert.ToBoolean(dt.Rows(0)("IsTransporterVehicle"))) Then
                                rbTransporterVeh.Checked = True
                                txtCarrName.Text = String.Empty
                                txtVehType.Text = String.Empty
                                If Convert.ToString(dt.Rows(0)("CONTRACT_TYPE")) <> "" Then
                                    If Convert.ToString(dt.Rows(0)("CONTRACT_TYPE")) = _contractTypeTripBased Then
                                        RadioButtonTrip.Checked = True
                                    ElseIf Convert.ToString(dt.Rows(0)("CONTRACT_TYPE")) = _contractTypeMonthlyBased Then
                                        RadioButtonMonthly.Checked = True
                                    ElseIf Convert.ToString(dt.Rows(0)("CONTRACT_TYPE")) = _contractTypeCourierBased Then
                                        RadioButtonCourier.Checked = True
                                        TextBoxZone.Text = dt.Rows(0)("ZONE")
                                        TextBoxTransportMode.Text = dt.Rows(0)("TRANSPORTER_MODE")
                                        TextBoxLenghtFT.Text = dt.Rows(0)("SIZE_LENGTH_FT")
                                        TextBoxLengthInch.Text = dt.Rows(0)("SIZE_LENGTH_INCH")
                                        TextBoxWidthFT.Text = dt.Rows(0)("SIZE_WIDTH_FT")
                                        TextBoxWidthInch.Text = dt.Rows(0)("SIZE_WIDTH_INCH")
                                        TextBoxHeightFT.Text = dt.Rows(0)("SIZE_HEIGHT_FT")
                                        TextBoxHeightInch.Text = dt.Rows(0)("SIZE_HEIGHT_INCH")
                                        TextBoxSize.Text = dt.Rows(0)("SIZE_CUBIC_FT")
                                        ctlWeight.Text = dt.Rows(0)("WEIGHT_KG")
                                    End If
                                End If
                            Else
                                rbCustVeh.Checked = True
                                txtCarrName.Text = Convert.ToString(dt.Rows(0)("CarrierName"))
                                txtVehType.Text = Convert.ToString(dt.Rows(0)("VehicleType"))
                            End If
                            txtTransporterCode.Text = Convert.ToString(dt.Rows(0)("Transporter_Code"))
                            txtVehicleCategory.Text = Convert.ToString(dt.Rows(0)("Vehicle_Category_Code"))
                            txtVehicleNo.Text = Convert.ToString(dt.Rows(0)("Vehicle_No"))

                            If Convert.ToString(dt.Rows(0)("adhocReqNo")).Length > 0 Then
                                rdoAdHoc.Checked = True
                                txtReqNo.Text = Convert.ToString(dt.Rows(0)("adhocReqNo"))
                            Else
                                rdoAdHoc.Checked = False
                                txtReqNo.Text = ""
                            End If

                            strSQL = "SELECT DocTypeTbl.RowNo RowNo, Document_Type InvoiceType, Document_No Doc_No, Document_Date DOC_DT" & _
                                    " , Customer_Code CUSTOMERCODE, Vendor_Code VendorCode, FrmLocation FromLocation, Vendor_WH_Code VendorWhCode, Doc_Qty Quantity, Doc_Value DocValue" & _
                                    " from Freight_Trip_Gen_Dtl FTD left join" & _
                                    " (" & _
                                    "   select 1 RowNo, 'INVOICE' InvoiceType" & _
                                    "   union" & _
                                    "   select 2 RowNo, 'RGP' InvoiceType" & _
                                    "   union" & _
                                    "   select 3 RowNo, 'NRGP' InvoiceType" & _
                                    " ) DocTypeTbl on FTD.Document_Type=DocTypeTbl.InvoiceType" & _
                                    " where FTD.Trip_No='" + strHelp(0) + "' and FTD.UNIT_CODE='" + gstrUNITID + "'"
                            dtDocDtl = SqlConnectionclass.GetDataTable(strSQL)
                            If Not IsNothing(dtDocDtl) AndAlso dtDocDtl.Rows.Count > 0 Then
                                fillGrid()
                            End If
                            InitializeForm(3)
                        End If
                    End If
                End If
            ElseIf sender Is BtnHelpTransporter Then
                Dim contractType As String = String.Empty
                If RadioButtonTrip.Checked Then
                    contractType = _contractTypeTripBased
                ElseIf RadioButtonMonthly.Checked Then
                    contractType = _contractTypeMonthlyBased
                ElseIf RadioButtonCourier.Checked Then
                    contractType = _contractTypeCourierBased
                End If
                strSQL = " SELECT VM.VENDOR_CODE, VM.VENDOR_NAME FROM VENDOR_MST VM WHERE VM.TRANSPORTER_FLAG=1 AND VM.ACTIVE_FLAG='A' AND VM.UNIT_CODE='" + gstrUNITID + "'" & _
                         " AND EXISTS(" & _
                         "     SELECT * FROM FREIGHT_CONTRACT_HDR FCH WHERE FCH.CONTRACT_TYPE='" & contractType & "' AND FCH.AUTHORIZED_STATUS='AUTHORIZED' AND FCH.UNIT_CODE=VM.UNIT_CODE AND FCH.TRANSPORTER_CODE=VM.VENDOR_CODE" & _
                         " )" & _
                         " ORDER BY VENDOR_CODE"

                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Transporter", 1, 0, txtTripNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 AndAlso Not IsNothing(strHelp(1)) Then
                        txtTransporterCode.Text = strHelp(0).Trim
                        txtTransporterName.Text = strHelp(1).Trim()
                        If mblnAllowTransporterfromMaster Then
                            txtVehicleNo.Text = ""
                        End If
                        clearGrid()
                    End If
                End If
            ElseIf sender Is BtnHelpVehCategory Then
                strSQL = " SELECT VEHICLE_CATEGORY_CODE, DESCRIPTION FROM FREIGHT_VEHICLE_VENDOR_CATEGORY_MST WHERE ISHOLD=0 AND UNIT_CODE='" + gstrUNITID + "' ORDER BY VEHICLE_CATEGORY_CODE"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Transporter", 1, 0, txtTripNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 AndAlso Not IsNothing(strHelp(1)) Then
                        txtVehicleCategory.Text = strHelp(0).Trim
                        clearGrid()
                        AddNewRow()
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtSplChars_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVehicleNo.KeyPress, txtTripNo.KeyPress, txtCarrName.KeyPress, txtVehType.KeyPress
        Try
            If e.KeyChar = "'" Then
                e.Handled = True
            End If
            If sender Is txtVehicleNo Then
                If Not ((e.KeyChar >= "A" And e.KeyChar <= "Z") Or (e.KeyChar >= "a" And e.KeyChar <= "z") Or (e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = "-" Or System.Text.Encoding.ASCII.GetBytes(e.KeyChar)(0) = 8) Then
                    e.Handled = True
                End If
            ElseIf sender Is txtCarrName Or sender Is txtVehType Then
                If Not ((e.KeyChar >= "A" And e.KeyChar <= "Z") Or (e.KeyChar >= "a" And e.KeyChar <= "z") Or (e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = "-" Or System.Text.Encoding.ASCII.GetBytes(e.KeyChar)(0) = 32 Or System.Text.Encoding.ASCII.GetBytes(e.KeyChar)(0) = 8) Then
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
        Try
            Dim dtData As DataTable
            Dim strSQL As String = String.Empty
            dtData = SqlConnectionclass.GetDataTable("SELECT TRIP_NO, CONVERT(VARCHAR(20),TRIP_DATE,106) TRIPDATE,ISNULL(TRANSPORTER_CODE,'') TRANSPORTER_CODE, ISNULL(VEHICLE_CATEGORY_CODE,'') VEHICLE_CATEGORY_CODE, VEHICLE_NO,IsTransporterVehicle" & _
                    " FROM FREIGHT_TRIP_GEN_HDR(NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND TRIP_NO='" + txtTripNo.Text.Trim() + "'")

            If Not IsNothing(dtData) AndAlso dtData.Rows.Count > 0 Then
                txtTripNo.Text = dtData.Rows(0)("Trip_No")
                dtpTripDt.Value = dtData.Rows(0)("TripDate")
                txtTransporterCode.Text = dtData.Rows(0)("Transporter_Code")
                txtVehicleCategory.Text = dtData.Rows(0)("Vehicle_Category_Code")
                txtVehicleNo.Text = dtData.Rows(0)("Vehicle_No")
                '101053952-starts
                If (Convert.ToBoolean(dtData.Rows(0)("IsTransporterVehicle"))) = True Then
                    rbTransporterVeh.Checked = True
                Else
                    rbCustVeh.Checked = True
                End If
                '101053952--ends
                strSQL = "SELECT DocTypeTbl.RowNo RowNo, Document_Type InvoiceType, Document_No Doc_No, Document_Date DOC_DT, Customer_Code CUSTOMERCODE" & _
                        " , Vendor_Code VendorCode, FrmLocation FromLocation, Vendor_WH_Code VendorWhCode, Doc_Qty Quantity, Doc_Value DocValue" & _
                        " from Freight_Trip_Gen_Dtl FTD left join" & _
                        " (" & _
                        "   select 1 RowNo, 'INVOICE' InvoiceType" & _
                        "   union" & _
                        "   select 2 RowNo, 'RGP' InvoiceType" & _
                        "   union" & _
                        "   select 3 RowNo, 'NRGP' InvoiceType" & _
                        ") DocTypeTbl on FTD.Document_Type=DocTypeTbl.InvoiceType" & _
                        " where FTD.Trip_No='" + dtData.Rows(0)("Trip_No") + "' and FTD.UNIT_CODE='" + gstrUNITID + "'"
                dtDocDtl = SqlConnectionclass.GetDataTable(strSQL)
                If Not IsNothing(dtDocDtl) AndAlso dtDocDtl.Rows.Count > 0 Then
                    fillGrid()
                End If
                InitializeForm(3)
            ElseIf Not String.IsNullOrEmpty(txtTripNo.Text) Then
                MessageBox.Show("Invalid Trip No", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtTripNo.Text = ""
                txtTripNo.Focus()
                InitializeForm(1)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTransporterCode.KeyDown, txtTripNo.KeyDown, txtVehicleCategory.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                If sender Is txtTripNo And cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    BtnHelp_Click(BtnHelpTrip, New EventArgs())
                ElseIf cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If sender Is txtTransporterCode Then
                        BtnHelp_Click(BtnHelpTransporter, New EventArgs())
                    ElseIf sender Is txtVehicleCategory Then
                        BtnHelp_Click(BtnHelpVehCategory, New EventArgs())
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fpSpreadDocSel_KeyPressEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles fpSpreadDocSel.KeyPressEvent
        Try
            If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If e.keyAscii = 14 Then    'Ctrl+N
                    AddNewRow()
                    GrpBoxSel.Enabled = Not (fpSpreadDocSel.MaxRows > 0)
                ElseIf e.keyAscii = 4 Then    'Ctrl+D
                    fpSpreadDocSel.Row = fpSpreadDocSel.ActiveRow
                    fpSpreadDocSel.Col = Enum_DocSel.Col_DocType
                    If Not IsNothing(dtDocDtl) AndAlso dtDocDtl.Select("RowNo=" + Convert.ToString(fpSpreadDocSel.CellTag)).Length > 0 Then
                        dtDocDtl.DefaultView.RowFilter = "RowNo not in ('" + Convert.ToString(fpSpreadDocSel.CellTag) + "')"
                        dtDocDtl = dtDocDtl.DefaultView.ToTable()
                        fillGrid()
                    Else
                        fpSpreadDocSel.Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                        fpSpreadDocSel.MaxRows = fpSpreadDocSel.MaxRows - 1
                    End If
                    GrpBoxSel.Enabled = Not (fpSpreadDocSel.MaxRows > 0)
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fpSpreadDocSel_KeyDownEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fpSpreadDocSel.KeyDownEvent
        Try
            Dim strSQL As String = String.Empty
            Dim strHelp As String()
            Dim strSelected As String = String.Empty
            Dim strDocNo As String = String.Empty
            Dim strFrmLoc As String = String.Empty
            Dim VendorCode As String = String.Empty
            Dim DocType As String = String.Empty
            Dim rNo As Integer = 1
            If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If e.keyCode = Keys.F1 AndAlso fpSpreadDocSel.ActiveRow > FPSpreadADO.CoordConstants.SpreadHeader Then
                    With ctlHelp
                        .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                        .ConnectAsUser = gstrCONNECTIONUSER
                        .ConnectThroughDSN = gstrCONNECTIONDSN
                        .ConnectWithPWD = gstrCONNECTIONPASSWORD
                    End With
                    With fpSpreadDocSel
                        If fpSpreadDocSel.ActiveCol = Enum_DocSel.Col_DocNo Then
                            fillDocNo()
                        ElseIf fpSpreadDocSel.ActiveCol = Enum_DocSel.Col_VendorWhCode Then
                            If RadioButtonMonthly.Checked Or RadioButtonCourier.Checked Then
                                Exit Sub
                            End If
                            .Row = .ActiveRow
                            rNo = .ActiveRow
                            .Col = Enum_DocSel.Col_CustCode
                            If Not String.IsNullOrEmpty(.Value.Trim()) Then
                                Return
                            End If
                            .Col = Enum_DocSel.Col_DocType
                            DocType = Convert.ToString(.Text)
                            .Col = Enum_DocSel.Col_DocNo
                            strDocNo = Convert.ToString(.Value)
                            .Col = Enum_DocSel.Col_FromLocation
                            strFrmLoc = Convert.ToString(.Value)
                            Dim rec() As DataRow = dtDocDtl.Select("InvoiceType='" + DocType + "' and Doc_No='" + strDocNo + "' and FromLocation='" + strFrmLoc + "'", "Quantity desc, DocValue desc")
                            If Not IsNothing(rec) AndAlso rec.Length > 0 Then
                                Dim itemArr() As Object
                                itemArr = rec(0).ItemArray
                                .Col = Enum_DocSel.Col_VendorWhCode
                                strSelected = .Value
                                .Col = Enum_DocSel.Col_VendorCode
                                strSQL = "SELECT WH_CODE [WAREHOUSECODE], ISNULL(ADDRESS2,'')+ISNULL(ADDRESS2,'') [ADDRESS], isnull(CITY,'') CITY, ISNULL(DIST,'') DISTRICT, isnull(STATE,'') STATE FROM FREIGHT_VENDOR_WHDEF_MST" & _
                                        " WHERE ACTIVE=1 AND VENDOR_CODE='" + Convert.ToString(.Value) + "' AND UNIT_CODE='" + gstrUNITID + "' "
                                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Warehouse(s)", 0, 0, strSelected)
                                If Not IsNothing(strHelp) AndAlso IsNothing(strHelp(1)) Then
                                    MessageBox.Show("Warehouse not defined for selected vendor.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    Return
                                End If
                                If Not IsNothing(strHelp) AndAlso strHelp.Length >= 5 Then
                                    If Not isContractExist(strHelp(0).Trim(), txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                        MessageBox.Show("Contract not defined for selected Vendor warehouse.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                        Return
                                    End If
                                    dtDocDtl.Rows(dtDocDtl.Rows.IndexOf(rec(0)))("VendorWhCode") = strHelp(0).Trim()
                                    dtDocDtl.AcceptChanges()
                                    .Row = rNo
                                    .Col = Enum_DocSel.Col_VendorWhCode
                                    .Value = strHelp(0).Trim()
                                End If
                                .Row = rNo
                                .Col = Enum_DocSel.Col_VendorWhCode
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            End If
                        End If
                    End With
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fpSpreadDocSel_ClickEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles fpSpreadDocSel.ClickEvent
        Try
            Dim strSql As String = String.Empty
            Dim dt As DataTable
            If fpSpreadDocSel.ActiveRow > FPSpreadADO.CoordConstants.SpreadHeader Then
                With fpSpreadDocSel
                    .Row = e.row
                    .Col = Enum_DocSel.Col_CustCode
                    If String.IsNullOrEmpty(.Value) Then
                        .Col = Enum_DocSel.Col_VendorCode
                        strSql = "select Vendor_name, (isnull(Office1_Address1,'')+isnull(Office1_Address2,'')) [Address], isnull(Office1_City,'') Office1_City, isnull(Office1_dist,'') Office1_dist, isnull(Office1_State,'') Office1_State " & _
                        " from vendor_mst where unit_code='" + gstrUNITID + "' and Vendor_code='" + .Value + "'"
                        dt = SqlConnectionclass.GetDataTable(strSql)
                        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                            txtCustName.Text = String.Empty
                            txtVendorCode.Text = dt.Rows(0)("Vendor_name")
                            txtAddress.Text = dt.Rows(0)("Address")
                            txtCity.Text = dt.Rows(0)("Office1_City")
                            txtDist.Text = dt.Rows(0)("Office1_dist")
                            txtState.Text = dt.Rows(0)("Office1_State")
                        End If
                    Else
                        strSql = "SELECT CUST_NAME, (ISNULL(OFFICE_ADDRESS1,'')+ISNULL(OFFICE_ADDRESS2,'')) [ADDRESS], isnull(OFFICE_CITY,'') OFFICE_CITY, isnull(OFFICE_DIST,'') OFFICE_DIST, isnull(OFFICE_STATE,'') OFFICE_STATE " & _
                        " FROM CUSTOMER_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" + .Value + "'"
                        dt = SqlConnectionclass.GetDataTable(strSql)
                        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                            txtCustName.Text = dt.Rows(0)("CUST_NAME")
                            txtVendorCode.Text = String.Empty
                            txtAddress.Text = dt.Rows(0)("Address")
                            txtCity.Text = dt.Rows(0)("OFFICE_CITY")
                            txtDist.Text = dt.Rows(0)("OFFICE_DIST")
                            txtState.Text = dt.Rows(0)("OFFICE_STATE")
                        End If
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fpSpreadDocSel_LeaveCell(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles fpSpreadDocSel.LeaveCell
        Try
            fpSpreadDocSel_ClickEvent(fpSpreadDocSel, New AxFPSpreadADO._DSpreadEvents_ClickEvent(e.newCol, e.newRow))
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dtpTripDt_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpTripDt.Validating
        Try
            clearGrid()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "Routines & Functions"

    ''' <param name="Form_Status_flag"> Form Initialization No </param>
    ''' <remarks>Initialize form</remarks>
    Private Sub InitializeForm(ByVal Form_Status_flag As Integer)
        Try
            If Form_Status_flag = 1 Then   'Page Load
                txtTripNo.Text = ""
                txtTripNo.Enabled = True
                txtTripNo.ReadOnly = False
                dtpTripDt.Enabled = False
                dtpTripDt.Value = GetServerDate()
                BtnHelpTransporter.Enabled = False
                txtTransporterCode.ReadOnly = True
                txtTransporterCode.Text = ""
                txtTransporterName.Text = String.Empty
                BtnHelpVehCategory.Enabled = False
                txtVehicleCategory.ReadOnly = True
                txtVehicleCategory.Text = ""
                txtVehicleNo.Text = String.Empty
                txtVehicleNo.ReadOnly = True
                txtCustName.Text = String.Empty
                txtCity.Text = String.Empty
                txtAddress.Text = String.Empty
                txtState.Text = String.Empty
                txtVendorCode.Text = String.Empty
                txtDist.Text = String.Empty
                BtnHelpTrip.Enabled = True
                txtCarrName.Text = String.Empty
                txtVehType.Text = String.Empty
                txtCarrName.ReadOnly = True
                txtVehType.ReadOnly = True
                GrpBoxSel.Enabled = False
                'BtnViewdoc.Enabled = False
                GroupBoxContract.Enabled = False
                TextBoxZone.Enabled = False
                ButtonZone.Enabled = False
                TextBoxTransportMode.Enabled = False
                ButtonTransportMode.Enabled = False
                TextBoxLenghtFT.Enabled = False
                TextBoxWidthFT.Enabled = False
                TextBoxHeightFT.Enabled = False
                TextBoxLengthInch.Enabled = False
                TextBoxWidthInch.Enabled = False
                TextBoxHeightInch.Enabled = False
                ctlWeight.Enabled = False
                TextBoxZone.Text = ""
                TextBoxTransportMode.Text = ""
                TextBoxLenghtFT.Text = ""
                TextBoxWidthFT.Text = ""
                TextBoxHeightFT.Text = ""
                TextBoxLengthInch.Text = ""
                TextBoxWidthInch.Text = ""
                TextBoxHeightInch.Text = ""
                ctlWeight.Text = ""

                txtReqNo.Text = String.Empty
                rdoAdHoc.Checked = False
                txtReqNo.Enabled = False

                cmdVehicleCodeHelp.Enabled = False
                AddDocSelectionGridColumn()
                AddColumnDocDtl()
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
                txtCarrName.Text = String.Empty
                txtVehType.Text = String.Empty
                txtCarrName.ReadOnly = False
                txtVehType.ReadOnly = False
                txtCustName.Text = String.Empty
                txtCity.Text = String.Empty
                txtAddress.Text = String.Empty
                txtVendorCode.Text = String.Empty
                txtDist.Text = String.Empty
                txtState.Text = String.Empty
                txtReqNo.Text = String.Empty
                dtpTripDt.Value = GetServerDateTime()
                GrpBoxSel.Enabled = True
                'If rdoAdHoc.Checked = True Then
                '    BtnViewdoc.Enabled = True
                'End If

                GroupBoxContract.Enabled = True
                cmdVehicleCodeHelp.Enabled = True
                clearGrid()
                txtVehicleCategory.Enabled = True
                BtnHelpVehCategory.Enabled = True
                BtnHelpTransporter.Enabled = True
                rbVeh_CheckedChanged(rbTransporterVeh, New EventArgs())
                dtpTripDt.Focus()
            ElseIf Form_Status_flag = 3 Then    'View Mode
                dtpTripDt.Enabled = False
                BtnHelpVehCategory.Enabled = False
                BtnHelpTrip.Enabled = True
                BtnHelpTransporter.Enabled = False
                txtVehicleNo.ReadOnly = True
                txtTripNo.ReadOnly = False
                txtTripNo.Enabled = True
                GrpBoxSel.Enabled = False
                txtCarrName.ReadOnly = True
                txtVehType.ReadOnly = True
                GroupBoxContract.Enabled = False
                TextBoxZone.Enabled = False
                ButtonZone.Enabled = False
                TextBoxTransportMode.Enabled = False
                ButtonTransportMode.Enabled = False
                TextBoxLenghtFT.Enabled = False
                TextBoxLengthInch.Enabled = False
                TextBoxWidthFT.Enabled = False
                TextBoxWidthInch.Enabled = False
                TextBoxHeightFT.Enabled = False
                TextBoxHeightInch.Enabled = False
                ctlWeight.Enabled = False
                cmdVehicleCodeHelp.Enabled = False
                If IsRecordExists(" SELECT Trip_No FROM Freight_Gate_Outward_Reg_Hdr WHERE Unit_code='" + gstrUNITID + "' and Trip_No='" + txtTripNo.Text.Trim() + "' AND CANCEL_YN = 'N'") Then
                    '101053952-starts
                    'If IsRecordExists(" SELECT Trip_No FROM Freight_Gate_Outward_Reg_Hdr WHERE Unit_code='" + gstrUNITID + "' and Trip_No='" + txtTripNo.Text.Trim() + "' AND CANCEL_YN = 'N'" & _
                    '                       " UNION " & _
                    '                       " SELECT Trip_No FROM Freight_Trip_Gen_Hdr " & _
                    '                       " WHERE Unit_Code='" + gstrUNITID + "' AND Trip_No='" + txtTripNo.Text.Trim() + "' and IsTransporterVehicle=0 and ISNULL(AuthStatus,'')='A'") Then
                    '101053952-ends
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                Else
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                End If
                cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
            ElseIf Form_Status_flag = 4 Then
                BtnHelpTrip.Enabled = False
                txtTripNo.Enabled = True
                txtTripNo.ReadOnly = True
                txtVehicleNo.ReadOnly = False
                cmdVehicleCodeHelp.Enabled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <remarks>To clear Database temporary tables data.</remarks>
    Private Function clearTmpTableData() As Boolean
        Dim flag As Boolean = True
        Dim strQry As String = String.Empty
        Try
            strQry = "delete from TMP_PR_PO_RFQ_VENDORS where IP_Address='" + gstrIpaddressWinSck + "'"
            SqlConnectionclass.ExecuteNonQuery(strQry)
            strQry = "delete from TMP_PR_PO_RFQ_ITEM_DETAIL where IP_Address='" + gstrIpaddressWinSck + "'"
            SqlConnectionclass.ExecuteNonQuery(strQry)
        Catch ex As Exception
            flag = False
        End Try
        Return flag
    End Function

    Private Function SaveDataAdhocInward() As Boolean
        Dim strSQL As String = ""
        Dim strTripNo As String = ""
        Dim oRDR As SqlDataReader

        Try

            If txtVehicleNo.Text = "" Or txtVehicleCategory.Text = "" Then
                MessageBox.Show("Please enter vehicle no and vehicle category.")
                Return False
            End If

            strTripNo = "TRIP" + Mid((100 + GetServerDate().Day()).ToString(), 2, 2) + Mid((100 + GetServerDate().Month()).ToString(), 2, 2) + GetServerDate.Year().ToString()


            strSQL = " SELECT '" + strTripNo + "' + RIGHT('0000' + CAST(ISNULL(MAX(CAST(REPLACE(TRIP_NO,'" + strTripNo + "','') AS INT)),0)+1 AS VARCHAR(20)),4)  from (" & _
                     " select Trip_No FROM FREIGHT_TRIP_GEN_HDR  WHERE UNIT_CODE='" + gstrUNITID + "' AND TRIP_NO LIKE '" + strTripNo + "%'  " & _
                     " Union" & _
                     " Select Trip_No FROM Freight_Gate_Inward  WHERE UNIT_CODE= '" + gstrUNITID + "' AND TRIP_NO LIKE '" + strTripNo + "%'  )a "

            oRDR = SqlConnectionclass.ExecuteReader(strSQL)
            If oRDR.HasRows Then
                oRDR.Read()
                strTripNo = oRDR(0).ToString()
            End If

            txtTripNo.Text = strTripNo

            strSQL = "Insert into Freight_Gate_Inward (AgainstDocument,RequestNo,Trip_No,Trip_Date,Trip_Value," & _
                    " Transporter_Code,Lorry_No,Vehicle_Category,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,UNIT_CODE)" & _
                    " values('Trip','" + txtReqNo.Text + "','" + txtTripNo.Text + "', getdate(), 0, '" + txtTransporterCode.Text + "', '" + txtVehicleNo.Text + "'," & _
                    " '" + txtVehicleCategory.Text + "', getdate(), '" + mP_User + "', getdate(), '" + mP_User + "', '" + gstrUNITID + "')"
            SqlConnectionclass.ExecuteNonQuery(strSQL)
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message())
            Return False
        End Try
    End Function


    ''' <returns>True for successful save.</returns>
    ''' <remarks></remarks>
    ''' 
    Private Function SaveData() As Boolean
        Try
            Dim strSql As String = String.Empty
            Dim contractType As String = String.Empty

            If rdoAdHoc.Checked = True Then
                contractType = _contractTypeTripBased
            End If

            If rdoAdHoc.Checked = True And rdoInward.Checked = True Then
                If SaveDataAdhocInward() = True Then
                    MessageBox.Show("Adhoc Inward Trip Generated.", ResolveResString(100), MessageBoxButtons.OK)
                    Return True
                Else
                    Return False
                End If
            End If
            If rdoAdHoc.Checked = True And rdoOutward.Checked = True Then
                If String.IsNullOrEmpty(txtVehicleNo.Text.Trim()) Then
                    MessageBox.Show("Vehicle No cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtVehicleNo.Focus()
                    Return False
                End If

                If String.IsNullOrEmpty(txtVehicleCategory.Text.Trim()) Then
                    MessageBox.Show("Vehicle Category cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtVehicleCategory.Focus()
                    Return False
                End If

            End If
            If rbTransporterVeh.Checked Then
                If RadioButtonTrip.Checked Then
                    contractType = _contractTypeTripBased
                ElseIf RadioButtonMonthly.Checked Then
                    contractType = _contractTypeMonthlyBased
                ElseIf RadioButtonCourier.Checked Then
                    contractType = _contractTypeCourierBased
                End If
            End If


            If rbCustVeh.Checked Then
                If String.IsNullOrEmpty(txtVehicleNo.Text.Trim()) Then
                    MessageBox.Show("Vehicle No cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtVehicleNo.Focus()
                    Return False
                End If
                If String.IsNullOrEmpty(txtCarrName.Text.Trim()) Then
                    MessageBox.Show("Carrier Name cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtCarrName.Focus()
                    Return False
                End If
                If String.IsNullOrEmpty(txtVehType.Text.Trim()) Then
                    MessageBox.Show("Vehicle Type cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtVehType.Focus()
                    Return False
                End If

            End If
            If rbTransporterVeh.Checked Then
                If String.IsNullOrEmpty(txtVehicleNo.Text.Trim()) Then
                    MessageBox.Show("Vehicle No cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtVehicleNo.Focus()
                    Return False
                End If

                If RadioButtonCourier.Checked Then
                    If Len(TextBoxZone.Text) = 0 Then
                        MessageBox.Show("Please enter Zone.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        TextBoxZone.Focus()
                        Return False
                    End If
                    If Len(TextBoxTransportMode.Text) = 0 Then
                        MessageBox.Show("Please enter Transport Mode.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        TextBoxTransportMode.Focus()
                        Return False
                    End If
                    If Val(TextBoxSize.Text) = 0 Then
                        MessageBox.Show("Please enter length, Width and Height.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        TextBoxLenghtFT.Focus()
                        Return False
                    End If
                    If Val(ctlWeight.Text) = 0 Then
                        MessageBox.Show("Please enter weight.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        ctlWeight.Focus()
                        Return False
                    End If
                End If
            End If

            fpSpreadDocSel.Row = fpSpreadDocSel.MaxRows
            fpSpreadDocSel.Col = Enum_DocSel.Col_DocNo
            If String.IsNullOrEmpty(fpSpreadDocSel.Value) Then
                MessageBox.Show("Document No cannot be empty.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                fpSpreadDocSel.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                Return False
            End If
            If dtDocDtl.Rows.Count = 0 Then
                MessageBox.Show("Select atleast One Document No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                fpSpreadDocSel.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                Return False
            End If

            If RadioButtonTrip.Checked Then
                fpSpreadDocSel.Row = fpSpreadDocSel.ActiveRow
                fpSpreadDocSel.Col = Enum_DocSel.Col_VendorCode


                strSql = " select top 1 1 from Vendor_mst where Unit_Code='" & gstrUNITID & "' and Vendor_Code = '" & fpSpreadDocSel.Value & "' "

                If IsRecordExists(strSql) = True Then
                    If dtDocDtl.Select("VendorCode<>'' and VendorWhCode=''").Length > 0 Then
                        MessageBox.Show("Vendor Warehouse code cannot be empty for vendors.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        fpSpreadDocSel.Row = dtDocDtl.Rows.IndexOf(dtDocDtl.Select("VendorCode<>'' and VendorWhCode=''")(0)) + 1
                        fpSpreadDocSel.Col = Enum_DocSel.Col_VendorWhCode
                        fpSpreadDocSel.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Return False
                    End If
                End If
            End If

            Dim ArrList As New ArrayList()
            With fpSpreadDocSel
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = Enum_DocSel.Col_DocNo
                    ArrList.Add(.Value)
                Next
            End With

            If MessageBox.Show("Are you sure, you want to save?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                Return False
            End If

            ' checking invoice no is exists    20240115
            If ArrList.Count > 0 Then
                Temp_InvocieNo_Insert(ArrList)
                If TripNo_IsExists_Again_InvoiceNo() Then
                    Return False
                End If
            End If

            Dim cmd As SqlCommand = New SqlCommand
            Dim strDocstring As String = ""
            With cmd
                .Connection = SqlConnectionclass.GetConnection()
                .Transaction = .Connection.BeginTransaction
                .CommandType = CommandType.StoredProcedure
                'added by priti on 09/05/2019 to avoid save multiple entry
                'If Len(txtTripNo.Text) = 0 Then
                '    If dtDocDtl IsNot Nothing AndAlso dtDocDtl.Rows.Count > 0 Then
                '        For Each Row As DataRow In dtDocDtl.Rows
                '            Dim strType As String = Convert.ToString(Row("InvoiceType"))
                '            If strType = "Invoice" Then
                '                Dim strDocNo = Convert.ToString(Row("DOC_NO"))
                '                Dim strDocExist = SqlConnectionclass.ExecuteScalar("select Document_No from FREIGHT_TRIP_GEN_DTL where Document_No='" & strDocNo & "' and Document_Type='Invoice' and unit_code='" & gstrUNITID & "'")
                '                If Len(strDocExist) > 0 OrElse strDocExist <> "" Then
                '                    strDocstring = strDocExist.ToString + "," + strDocstring.ToString
                '                End If
                '            End If
                '        Next
                '        If Len(strDocstring) > 0 OrElse strDocstring <> "" Then
                '            strDocstring = strDocstring.Substring(0, strDocstring.Length - 1)
                '            MessageBox.Show("Trip No has been already created for Invoice " & strDocstring & " ", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                '            .Transaction.Rollback()
                '            Return False
                '        End If
                '    End If
                'End If
                'Priti 09/05/2019 code ends here
                If Len(txtTripNo.Text.Trim) = 0 Then
                    dtpTripDt.Value = GetServerDateTime()
                End If

                .CommandText = "USP_FREIGHT_TRIP_GENERATION"
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
                .Parameters.AddWithValue("@p_TripTbl", dtDocDtl)
                .Parameters.Add(New SqlParameter("@P_Error", SqlDbType.VarChar, 500, ParameterDirection.Output, True, 0, 0, "", DataRowVersion.Default, ""))
                .Parameters.AddWithValue("@p_isTransport", IIf(rbTransporterVeh.Checked Or rdoAdHoc.Checked, 1, 0))
                .Parameters.AddWithValue("@p_CarrName", txtCarrName.Text.Trim())
                .Parameters.AddWithValue("@p_VehicleType", txtVehType.Text.Trim())
                .Parameters.AddWithValue("@p_AuthBy", "00000000")
                If rbTransporterVeh.Checked Or rdoAdHoc.Checked = True Then
                    .Parameters.AddWithValue("@CONTRACT_TYPE", contractType)
                    If RadioButtonCourier.Checked Then
                        .Parameters.AddWithValue("@SIZE_LENGTH_FT", Val(TextBoxLenghtFT.Text))
                        .Parameters.AddWithValue("@SIZE_LENGTH_INCH", Val(TextBoxLengthInch.Text))
                        .Parameters.AddWithValue("@SIZE_WIDTH_FT", Val(TextBoxWidthFT.Text))
                        .Parameters.AddWithValue("@SIZE_WIDTH_INCH", Val(TextBoxWidthInch.Text))
                        .Parameters.AddWithValue("@SIZE_HEIGHT_FT", Val(TextBoxHeightFT.Text))
                        .Parameters.AddWithValue("@SIZE_HEIGHT_INCH", Val(TextBoxHeightInch.Text))
                        .Parameters.AddWithValue("@SIZE_CUBIC_FT", Val(TextBoxSize.Text))
                        .Parameters.AddWithValue("@WEIGHT_KG", Val(ctlWeight.Text))
                        .Parameters.AddWithValue("@ZONE", TextBoxZone.Text.Trim())
                        .Parameters.AddWithValue("@TRANSPORTER_MODE", TextBoxTransportMode.Text.Trim())
                    End If
                End If
                .ExecuteNonQuery()
                If String.IsNullOrEmpty(.Parameters("@P_Error").Value) Then

                    txtTripNo.Text = .Parameters("@P_TripNo").Value.ToString()

                    If rdoAdHoc.Checked = True And txtReqNo.Text.Length > 0 Then
                        strSql = "Update FREIGHT_TRIP_GEN_HDR set adhocReqNo = '" + txtReqNo.Text + "' where unit_code = '" + gstrUNITID + "' and Trip_No = '" + txtTripNo.Text + "'"
                        .Parameters.Clear()
                        .CommandType = CommandType.Text
                        .CommandText = strSql
                        .ExecuteNonQuery()
                    End If

                    .Transaction.Commit()

                    MessageBox.Show("Data saved successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return True
                    'ElseIf Convert.ToString(.Parameters("@P_Error").Value).Contains("~") Then
                    '    .Transaction.Commit()
                    '    txtTripNo.Text = .Parameters("@P_TripNo").Value.ToString()
                    '    MessageBox.Show("Data saved successfully, But Email cannot be sent to authorised person due to following reason." + vbCrLf + Convert.ToString(.Parameters("@P_Error").Value).Replace("~", ""), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    '    Return True
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

    Private Function TripNo_IsExists_Again_InvoiceNo() As Boolean
        Try
            Dim sqlCmd As New SqlCommand
            With sqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_FREIGHT_TRIP_IS_EXISTS"
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                .Parameters.Add(New SqlParameter("@EXISTS_INVOICENO", SqlDbType.VarChar, 4000, ParameterDirection.Output, True, 0, 0, "", DataRowVersion.Default, ""))
                .ExecuteNonQuery()
                If Convert.ToString(.Parameters("@EXISTS_INVOICENO").Value).Trim().Length > 0 Then
                    MessageBox.Show(Convert.ToString(.Parameters("@EXISTS_INVOICENO").Value).Trim().Replace(",", Environment.NewLine).Replace("#", ""), "Trip has been already created for these invoices", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return True
                Else
                    Return False
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            Return True
        End Try
    End Function

    Private Sub Temp_InvocieNo_Insert(ByVal InvoiceNoList As ArrayList)
        Try
            Temp_InvocieNo_Remove()
            For Each InvoiceNo As String In InvoiceNoList
                Dim sqlcmd As New SqlCommand
                With sqlcmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .Transaction = .Connection.BeginTransaction
                    .CommandType = CommandType.Text
                    .CommandTimeout = 30
                    .CommandText = "INSERT INTO TEMP_TRIPGEN_AGAINS_INVOICENO(INVOICE_NO, UNIT_CODE, IP_ADDRESS) VALUES(@INVOICE_NO, @UNIT_CODE, @IP_ADDRESS)"
                    .Parameters.AddWithValue("@INVOICE_NO", InvoiceNo)
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                    Try
                        .ExecuteNonQuery()
                        .Transaction.Commit()
                    Catch ex As Exception
                        .Transaction.Rollback()
                        Continue For
                    End Try
                End With
            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub Temp_InvocieNo_Remove()
        Try
            Dim sqlcmd As New SqlCommand
            With sqlcmd
                .Connection = SqlConnectionclass.GetConnection()
                .Transaction = .Connection.BeginTransaction
                .CommandType = CommandType.Text
                .CommandTimeout = 30
                .CommandText = "DELETE FROM TEMP_TRIPGEN_AGAINS_INVOICENO WHERE UNIT_CODE = @UNIT_CODE AND IP_ADDRESS = @IP_ADDRESS;"
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                Try
                    .ExecuteNonQuery()
                    .Transaction.Commit()
                Catch ex As Exception
                    .Transaction.Rollback()
                End Try
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <returns>TRUE IF RECORD DELETED SUCCESSFULLY.</returns>
    ''' <remarks>To Delete RFQ in Delete press.</remarks>
    Private Function DeleteData() As Boolean
        Try
            Dim sqlCmd As New SqlCommand
            With sqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.StoredProcedure
                .Transaction = .Connection.BeginTransaction
                .CommandTimeout = 0
                .CommandText = "USP_FREIGHT_TRIP_GENERATION"
                .Parameters.AddWithValue("@p_UnitCode", gstrUNITID)
                .Parameters.AddWithValue("@p_TranType", "DELETE")
                .Parameters.AddWithValue("@p_UserId", mP_User)
                .Parameters.Add(New SqlParameter("@P_TripNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                .Parameters("@P_TripNo").Value = txtTripNo.Text.Trim
                .Parameters.Add(New SqlParameter("@P_Error", SqlDbType.NChar, 500, ParameterDirection.Output, True, 0, 0, "", DataRowVersion.Default, ""))
                .ExecuteNonQuery()
                If Not String.IsNullOrEmpty(Convert.ToString(.Parameters("@P_Error").Value).Trim) Then
                    .Transaction.Rollback()
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@P_Error").Value).Trim, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                    Return False
                Else
                    .Transaction.Commit()
                    MessageBox.Show("Trip No " + Convert.ToString(.Parameters("@P_TripNo").Value).Trim + " deleted Successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return True
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return False
    End Function

    Private Sub clearGrid()
        Try
            If (Not IsNothing(dtDocDtl)) Then
                fpSpreadDocSel.MaxRows = 0
                dtDocDtl.Rows.Clear()
                dtDocDtl.AcceptChanges()
                If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW And Not String.IsNullOrEmpty(txtTransporterCode.Text.Trim()) And Not String.IsNullOrEmpty(txtVehicleCategory.Text.Trim()) And (cmdVehicleCodeHelp.Enabled = True And txtVehicleNo.Text <> "") Then
                    AddNewRow()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub AddDocSelectionGridColumn()
        Try
            With fpSpreadDocSel
                .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
                .MaxRows = 0
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .set_RowHeight(.Row, 20)

                .MaxCols = Enum_DocSel.Col_DocType
                .Col = Enum_DocSel.Col_DocType
                .Value = "Doc. Type"
                .set_ColWidth(Enum_DocSel.Col_DocType, 10)

                .MaxCols = Enum_DocSel.Col_DocNo
                .Col = Enum_DocSel.Col_DocNo
                .Value = "Document No" + vbCrLf + "[F1]"
                .set_ColWidth(Enum_DocSel.Col_DocNo, 10)

                .MaxCols = Enum_DocSel.Col_DocDt
                .Col = Enum_DocSel.Col_DocDt
                .Value = "Document Date"
                .set_ColWidth(Enum_DocSel.Col_DocDt, 10)

                .MaxCols = Enum_DocSel.Col_CustCode
                .Col = Enum_DocSel.Col_CustCode
                .Value = "Customer Code"
                .set_ColWidth(Enum_DocSel.Col_CustCode, 10)

                .MaxCols = Enum_DocSel.Col_VendorCode
                .Col = Enum_DocSel.Col_VendorCode
                .Value = "Vendor Code"
                .set_ColWidth(Enum_DocSel.Col_VendorCode, 10)

                .MaxCols = Enum_DocSel.Col_FromLocation
                .Col = Enum_DocSel.Col_FromLocation
                .Value = "From Location"
                .set_ColWidth(Enum_DocSel.Col_FromLocation, 8)

                .MaxCols = Enum_DocSel.Col_VendorWhCode
                .Col = Enum_DocSel.Col_VendorWhCode
                .Value = "Vendor Warehouse Code" + vbCrLf + "[F1]"
                .set_ColWidth(Enum_DocSel.Col_VendorWhCode, 16)

                .MaxCols = Enum_DocSel.Col_Qty
                .Col = Enum_DocSel.Col_Qty
                .Value = "Qty."
                .set_ColWidth(Enum_DocSel.Col_Qty, 7)

                .MaxCols = Enum_DocSel.Col_DocValue
                .Col = Enum_DocSel.Col_DocValue
                .Value = "Doc. Value"
                .set_ColWidth(Enum_DocSel.Col_DocValue, 8)

                .BlockMode = True
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Row2 = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = Enum_DocSel.Col_DocType
                .Col2 = Enum_DocSel.Col_DocValue
                .Lock = True
                .BlockMode = False
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub AddNewRow()
        Try
            Dim strComboLst As String = String.Empty
            If String.IsNullOrEmpty(txtTransporterCode.Text.Trim()) And rbTransporterVeh.Checked Then
                MessageBox.Show("Transporter Code cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtTransporterCode.Focus()
                Return
            ElseIf String.IsNullOrEmpty(txtVehicleCategory.Text.Trim()) And rbTransporterVeh.Checked Then
                If RadioButtonTrip.Checked Or RadioButtonMonthly.Checked Then
                    MessageBox.Show("Vehicle Category cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtVehicleCategory.Focus()
                    Return
                End If
            End If
            If Not IsNothing(dtDocDtl) AndAlso dtDocDtl.Rows.Count > 0 Then
                If dtDocDtl.Select("InvoiceType='Invoice'").Length <= 0 Then
                    strComboLst = strComboLst & "Invoice" & Chr(9)
                End If
                If dtDocDtl.Select("InvoiceType='RGP'").Length <= 0 Then
                    strComboLst = strComboLst & "RGP" & Chr(9)
                End If
                If dtDocDtl.Select("InvoiceType='NRGP'").Length <= 0 Then
                    strComboLst = strComboLst & "NRGP" & Chr(9)
                End If
            Else
                strComboLst = "Invoice" & Chr(9) & "RGP" & Chr(9) & "NRGP"
            End If
            If String.IsNullOrEmpty(strComboLst) Then
                MessageBox.Show("All the document type has been added in the grid. Kindly add new document no in existing document type.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Return
            End If
            With fpSpreadDocSel
                .Row = .MaxRows
                .Col = Enum_DocSel.Col_DocNo
                If String.IsNullOrEmpty(.Value) Then
                    Return
                End If
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = Enum_DocSel.Col_DocType
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox

                .TypeComboBoxList = strComboLst
                .TypeComboBoxEditable = 0
                .TypeComboBoxIndex = 1
                .Value = "Invoice"
                If Not IsNothing(dtDocDtl) AndAlso dtDocDtl.Rows.Count > 0 Then
                    .CellTag = Integer.Parse(dtDocDtl.Compute("Max(RowNo)", "")) + 1
                Else
                    .CellTag = 1
                End If

                .Col = Enum_DocSel.Col_DocNo
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Value = ""

                .Col = Enum_DocSel.Col_DocDt
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Value = ""

                .Col = Enum_DocSel.Col_CustCode
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Value = ""

                .Col = Enum_DocSel.Col_VendorCode
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Value = ""

                .Col = Enum_DocSel.Col_VendorWhCode
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Value = ""

                .Col = Enum_DocSel.Col_Qty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Value = ""

                .Col = Enum_DocSel.Col_DocValue
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Value = ""

                .Col = Enum_DocSel.Col_DocNo
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub LockUnlockGrid()
        Try
            With fpSpreadDocSel
                If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = Enum_DocSel.Col_DocType
                    .Col2 = Enum_DocSel.Col_DocValue
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                Else
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .Col = Enum_DocSel.Col_DocType
                    .Col2 = Enum_DocSel.Col_DocType
                    .Lock = False
                    .Col = Enum_DocSel.Col_VendorWhCode
                    .Col2 = Enum_DocSel.Col_VendorWhCode
                    .Lock = False
                    .BlockMode = False
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub AddColumnDocDtl()
        Try
            If IsNothing(dtDocDtl) Then
                dtDocDtl = New DataTable()
            End If
            If dtDocDtl.Columns.Count <= 0 Then
                dtDocDtl.Columns.Add(New DataColumn("RowNo"))
                dtDocDtl.Columns.Add(New DataColumn("InvoiceType"))
                dtDocDtl.Columns.Add(New DataColumn("Doc_No"))
                dtDocDtl.Columns.Add(New DataColumn("Doc_Dt"))
                dtDocDtl.Columns.Add(New DataColumn("CustomerCode"))
                dtDocDtl.Columns.Add(New DataColumn("VendorCode"))
                dtDocDtl.Columns.Add(New DataColumn("FromLocation"))
                dtDocDtl.Columns.Add(New DataColumn("VendorWhCode"))
                dtDocDtl.Columns.Add(New DataColumn("Quantity", System.Type.GetType("System.Double")))
                dtDocDtl.Columns.Add(New DataColumn("DocValue", System.Type.GetType("System.Double")))
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fillGrid()
        Try
            Dim StartRow, EndRow As Integer
            CustomerCode = String.Empty
            CustomerCodes.Clear()
            If Not IsNothing(dtDocDtl) Then
                With fpSpreadDocSel
                    .MaxRows = 0
                    .MaxRows = .MaxRows + 1
                    dtDocDtl.DefaultView.Sort = "RowNo, Doc_No"
                    For Each dr As DataRow In dtDocDtl.DefaultView().ToTable(True, "RowNo", "InvoiceType").Rows
                        .Row = .MaxRows
                        StartRow = .Row
                        For Each rec As DataRow In dtDocDtl.Select("RowNo=" + Convert.ToString(dr("RowNo")) + " and InvoiceType='" + Convert.ToString(dr("InvoiceType")) + "'", "RowNo, Doc_No")
                            .Col = Enum_DocSel.Col_DocType
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            '.TypeComboBoxList = "Invoice" & Chr(9) & "RGP" & Chr(9) & "NRGP"
                            '.TypeComboBoxEditable = 0
                            '.TypeComboBoxIndex = 1
                            .Value = Convert.ToString(dr("InvoiceType"))
                            .CellTag = dr("RowNo")

                            If .Text.ToUpper() = "INVOICE" Then
                                .Col = Enum_DocSel.Col_DocValue
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .Value = dtDocDtl.DefaultView.ToTable(True, "RowNo", "InvoiceType", "Doc_No", "Doc_Dt", "FromLocation", "DocValue").Compute("SUM(DocValue)", "RowNo=" + Convert.ToString(dr("RowNo")) + " and InvoiceType='" + Convert.ToString(dr("InvoiceType")) + "'")
                            Else
                                .Col = Enum_DocSel.Col_DocValue
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .Value = "0"
                                .Text = " "
                            End If

                            .Col = Enum_DocSel.Col_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .Value = dtDocDtl.DefaultView.ToTable(True, "RowNo", "InvoiceType", "Doc_No", "Doc_Dt", "FromLocation", "Quantity").Compute("SUM(Quantity)", "RowNo=" + Convert.ToString(dr("RowNo")) + " and InvoiceType='" + Convert.ToString(dr("InvoiceType")) + "'")

                            .Col = Enum_DocSel.Col_DocNo
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .Value = rec("Doc_No")
                            .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted

                            .Col = Enum_DocSel.Col_DocDt
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .Value = rec("Doc_dt")
                            .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted

                            .Col = Enum_DocSel.Col_CustCode
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .Value = rec("CustomerCode")
                            .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted
                            CustomerCode = rec("CustomerCode").ToString()
                            CustomerCodes.Add(rec("CustomerCode").ToString())

                            .Col = Enum_DocSel.Col_VendorCode
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .Value = rec("VendorCode")
                            .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted

                            .Col = Enum_DocSel.Col_FromLocation
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .Value = rec("FromLocation")
                            .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted

                            .Col = Enum_DocSel.Col_VendorWhCode
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .Value = rec("VendorWhCode")
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            EndRow = .Row - 1
                        Next

                        .Row = StartRow
                        .Row2 = EndRow
                        .Col = Enum_DocSel.Col_Qty
                        .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                        .Col = Enum_DocSel.Col_DocValue
                        .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                        .Col = Enum_DocSel.Col_DocType
                        .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    Next
                    .MaxRows = .MaxRows - 1
                End With
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
            Dim TripAfterEwayBillDate As Integer = 0
            Dim errMsg As String = String.Empty
            Dim strEwayBillType As String = ""
            strSQL = String.Empty
            strSelectedRec = String.Empty
            If fpSpreadDocSel.MaxRows <= 0 Then
                Return
            End If
            If rbTransporterVeh.Checked Then
                If RadioButtonCourier.Checked Then
                    If Len(TextBoxZone.Text) = 0 Then
                        MessageBox.Show("Please enter Zone.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        TextBoxZone.Focus()
                        Return
                    End If
                    If Len(TextBoxTransportMode.Text) = 0 Then
                        MessageBox.Show("Please enter Transport Mode.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        TextBoxTransportMode.Focus()
                        Return
                    End If
                    If Val(TextBoxSize.Text) = 0 Then
                        MessageBox.Show("Please enter length, Width and Height.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        TextBoxLenghtFT.Focus()
                        Return
                    End If
                    If Val(ctlWeight.Text) = 0 Then
                        MessageBox.Show("Please enter weight.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        ctlWeight.Focus()
                        Return
                    End If
                End If
            End If
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            noOfDays = SqlConnectionclass.ExecuteScalar("SELECT DaysBehindCurDtForFreightTripGenDocNo from POConfig_Mst where UNIT_CODE='" + gstrUNITID + "'")
            'noOfDays = 365
            With fpSpreadDocSel
                .Row = .ActiveRow
                .Col = Enum_DocSel.Col_DocType
                rNo = .ActiveRow
                For Each row As DataRow In dtDocDtl.Select("RowNo=" + Convert.ToString(.CellTag))
                    strSelectedRec = strSelectedRec + row("Doc_No").ToString() + "|"
                Next
                If .Text.ToUpper() = "INVOICE" Then
                    Dim intCount As Integer = 0
                    Dim intCount2 As Integer = 0
                    fillInvoiceDocInTmp()

                    '  Added by priti for Eway bill validation start here
                    TripAfterEwayBillDate = SqlConnectionclass.ExecuteScalar("SELECT isnull(TripAfterEwayBillDate,0) from POConfig_Mst where UNIT_CODE='" + gstrUNITID + "'")
                    If TripAfterEwayBillDate = 1 Then
                        strEwayBillType = "EwayBillType as EwayBillType,"
                        intCount = 9
                        intCount2 = 10
                    Else
                        intCount = 8
                        intCount2 = 9
                    End If
                    ' Added by priti for Eway bill validation ends here

                    strSQL = "SELECT DOC_NO, CONVERT(VARCHAR(20), DOC_DT, 106) DOC_DT, INVOICE_TYPE," & strEwayBillType & " CUSTOMERCODE, VENDORCODE, CUST_NAME, DOCVALUE, QUANTITY   " & _
                            " FROM TMP_FREIGHT_TRIP_SELECT_DOC WHERE UNIT_CODE='" + gstrUNITID + "' AND IP_ADDRESS='" + gstrIpaddressWinSck + "' and DOCTYPE = 'INVOICE'  "
                    strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Document(s)", 0, 1, strSelectedRec)

                    If Not IsNothing(strHelp) AndAlso strHelp.Length < intCount Then
                        MessageBox.Show("No record found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return
                    ElseIf Not IsNothing(strHelp) AndAlso strHelp.Length = intCount Then
                        MessageBox.Show("Select atleast one document.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return
                    End If
                    If Not IsNothing(strHelp) AndAlso strHelp.Length = intCount2 Then
                        ' Added by priti for Eway bill validation start here
                        If TripAfterEwayBillDate = 1 Then
                            Dim intEWayATypeCount = SqlConnectionclass.ExecuteScalar("select count(*) from TMP_FREIGHT_TRIP_SELECT_DOC where DOC_NO in ('" + strHelp(intCount).Replace("|", "','") + "') and UNIT_CODE='" + gstrUNITID + "' and IP_ADDRESS='" + gstrIpaddressWinSck + "' and EwayBillType='A'")
                            If intEWayATypeCount > 0 Then
                                MessageBox.Show("Please select only EwayBill B Type invoices.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                clearGrid()
                                AddNewRow()
                                Return
                            End If
                        End If
                        ' Added by priti for Eway bill validation ends here
                        .Row = rNo
                        .Col = Enum_DocSel.Col_DocType
                        strSQL = "SELECT '" + Convert.ToString(.CellTag) + "' RowNo, '" + Convert.ToString(.Text) + "' InvoiceType, SD.DOC_NO, CONVERT(VARCHAR(20),SCD.INVOICE_DATE,106) DOC_DT, isnull((SELECT TOP 1 CUSTOMER_CODE FROM CUSTOMER_MST CM WHERE CM.UNIT_CODE =SCD.UNIT_CODE AND CM.CUSTOMER_CODE =SCD.ACCOUNT_CODE),'') CUSTOMERCODE" & _
                       " , isnull((SELECT TOP 1 VENDOR_CODE FROM VENDOR_MST VM WHERE VM.UNIT_CODE =SCD.UNIT_CODE AND VM.VENDOR_CODE =SCD.ACCOUNT_CODE),'') VENDORCODE" & _
                       " , '' FromLocation, '' VendorWhCode, SUM(ISNULL(SD.SALES_QUANTITY,0)) QUANTITY, SCD.TOTAL_AMOUNT [DOCVALUE]" & _
                       " FROM SALES_DTL(NOLOCK) SD INNER JOIN SALESCHALLAN_DTL(NOLOCK) SCD " & _
                       " ON SD.UNIT_CODE=SCD.UNIT_CODE AND SD.DOC_NO=SCD.DOC_NO" & _
                       " WHERE SD.UNIT_CODE='" + gstrUNITID + "' AND SCD.BILL_FLAG=1 AND SCD.CANCEL_FLAG=0" & _
                       " AND SCD.DOC_NO IN ('" + strHelp(intCount).Replace("|", "','") + "')" & _
                       " GROUP BY SD.DOC_NO, SCD.INVOICE_DATE, SCD.INVOICE_TYPE, SCD.ACCOUNT_CODE, SCD.CUST_NAME, SCD.TOTAL_AMOUNT,SCD.UNIT_CODE"
                        Dim dt As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                            For Each row As DataRow In dtDocDtl.DefaultView.ToTable(True, "InvoiceType", "RowNo", "Doc_No", "FromLocation").Select("RowNo in ('" + Convert.ToString(.CellTag) + "')")
                                If dt.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'").Length > 0 Then
                                    dt.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'")(0).Delete()
                                Else
                                    For Each r As DataRow In dtDocDtl.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'")
                                        r.Delete()
                                    Next
                                End If
                            Next
                            dt.AcceptChanges()
                            dtDocDtl.AcceptChanges()
                            Dim insertRow As DataRow
                            errMsg = String.Empty
                            For Each row As DataRow In dt.Rows
                                If Not String.IsNullOrEmpty(row("CUSTOMERCODE").ToString().Trim()) And Not isContractExist(row("CUSTOMERCODE").ToString().Trim(), txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                    If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                        errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                    End If
                                    Continue For
                                End If
                                If rbTransporterVeh.Checked Then
                                    If RadioButtonMonthly.Checked Then
                                        If Not isContractExistMonthly(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                    End If
                                    If RadioButtonCourier.Checked Then
                                        If Not isContractExistSize(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                        If Not isContractExistWeight(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                    End If
                                End If
                                insertRow = dtDocDtl.NewRow
                                insertRow("RowNo") = row("RowNo")
                                insertRow("InvoiceType") = row("InvoiceType")
                                insertRow("Doc_No") = row("DOC_NO")
                                insertRow("Doc_Dt") = row("DOC_DT")
                                insertRow("CustomerCode") = row("CUSTOMERCODE")
                                insertRow("VendorCode") = row("VENDORCODE")
                                insertRow("FromLocation") = row("FromLocation")
                                insertRow("VendorWhCode") = row("VendorWhCode")
                                insertRow("Quantity") = row("QUANTITY")
                                insertRow("DocValue") = row("DOCVALUE")
                                dtDocDtl.Rows.Add(insertRow)
                            Next
                            If dtDocDtl.Rows.Count > 0 Then
                                fillGrid()
                            End If
                            If Not String.IsNullOrEmpty(errMsg) Then
                                If rbTransporterVeh.Checked Then
                                    If RadioButtonCourier.Checked Then
                                        errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + "." + errMsg
                                    Else
                                        errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + " and Vehicle Category-" + txtVehicleCategory.Text.Trim() + "." + errMsg
                                    End If
                                Else
                                    errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + " and Vehicle Category-" + txtVehicleCategory.Text.Trim() + "." + errMsg
                                End If
                                MessageBox.Show(errMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            End If
                            .Row = rNo
                            .Col = Enum_DocSel.Col_DocNo
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    End If
                ElseIf .Text = "NRGP" Then      '20240118
                    fillRGPDocInTmp(.Text)
                    strSQL = "SELECT DOC_NO, FROM_LOCATION,CONVERT(VARCHAR(20), DOC_DT, 106) DOC_DT, " & strEwayBillType & " CUSTOMERCODE, VENDORCODE, CUST_NAME as [NAME], QUANTITY as TOTALQUANTITY   " & _
                      " FROM TMP_FREIGHT_TRIP_SELECT_DOC WHERE UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "' and DOCTYPE = '" & .Text & "'  "

                    'strSQL = "select CONVERT(VARCHAR(20),RH.Doc_No) DOC_NO, RH.From_Location, convert(varchar(20),RH.RGP_Date,106) Doc_Dt, isnull(CM.Customer_Code,'') CustomerCode, isnull(VM.Vendor_Code,'') VendorCode, case when RH.Location_To_Type='V' then VM.Vendor_name else CM.Cust_Name end [Name], sum(isnull(RD.Actual_Quantity,0)) TotalQuantity" & _
                    '        " from RGP_Hdr RH " & _
                    '        " inner join RGP_Dtl RD on RH.UNIT_CODE=RD.UNIT_CODE and RH.Doc_No=RD.Doc_No and RD.From_Location=RH.From_Location and RH.Doc_Type=RD.Doc_Type" & _
                    '        " left join Vendor_mst VM on RH.To_location=VM.Vendor_loc and RH.UNIT_CODE=VM.UNIT_CODE and RH.Location_To_Type='V'" & _
                    '        " left join Customer_mst CM on RH.To_location=CM.Cust_Location and RH.UNIT_CODE=CM.UNIT_CODE and RH.Location_To_Type='C'" & _
                    '        " where RH.UNIT_CODE='" + gstrUNITID + "' and RH.NRGP_Cancelled=0 and isnull(RH.Authorized_Code,'')<>'' and RH.Location_To_Type in ('V','C')" & _
                    '        " and RD.Actual_Quantity>0 and RH.RGP_Date between '" + dtpTripDt.Value.AddDays(-1 * noOfDays).ToString("dd MMM yyyy") + "' and '" + dtpTripDt.Value.ToString("dd MMM yyyy") + "'" & _
                    '        " and RH.Doc_Type=23" & _
                    '        " and not exists" & _
                    '        " (" & _
                    '        "       select DOCUMENT_TYPE, DOCUMENT_NO, FRMLOCATION from" & _
                    '        "       (" & _
                    '        "               SELECT DISTINCT CASE WHEN ISNULL(TBL.TRIP_NO,'')<>'' THEN ISNULL(TBL.TRIP_DOC_TYPE,'') ELSE FTD.DOCUMENT_TYPE END DOCUMENT_TYPE       " & _
                    '        "               , CASE WHEN ISNULL(TBL.TRIP_NO,'')<>'' THEN ISNULL(TBL.TRIP_DOC_NO,'') ELSE FTD.DOCUMENT_NO END  DOCUMENT_NO" & _
                    '        "               , CASE WHEN ISNULL(TBL.TRIP_NO,'')<>'' THEN ISNULL(TBL.FRMLOCATION,'') ELSE FTD.FRMLOCATION END  FRMLOCATION " & _
                    '        "               FROM FREIGHT_TRIP_GEN_DTL FTD LEFT JOIN        " & _
                    '        "               (" & _
                    '        "                   SELECT FGH.UNIT_CODE, FGH.TRIP_NO, FGTD.TRIP_DOC_NO, FGTD.TRIP_DOC_TYPE, FGTD.CUST_VEND_CODE, FGTD.FRMLOCATION " & _
                    '        "                   FROM FREIGHT_GATE_OUTWARD_REG_HDR FGH" & _
                    '        "                   INNER JOIN FREIGHT_GATE_OUTWARD_TRIP_DOC_DTL FGTD ON FGH.DOC_NO=FGTD.DOC_NO AND FGH.UNIT_CODE=FGTD.UNIT_CODE" & _
                    '        "                   WHERE FGH.UNIT_CODE='" + gstrUNITID + "'" & _
                    '        "               ) TBL " & _
                    '        "               ON FTD.UNIT_CODE=TBL.UNIT_CODE AND FTD.TRIP_NO=TBL.TRIP_NO AND FTD.DOCUMENT_NO=TBL.TRIP_DOC_NO AND FTD.DOCUMENT_TYPE=TBL.TRIP_DOC_TYPE AND FTD.FRMLOCATION=TBL.FRMLOCATION" & _
                    '        "               WHERE FTD.UNIT_CODE='" + gstrUNITID + "' and FTD.Ent_dt > = '01 Apr 2022'     " & _
                    '        "               union" & _
                    '        "               SELECT AGAINSTDOCUMENT DOCUMENT_TYPE, AGAINSTDOCUMENTNUMBER DOCUMENT_NO, LOCATION_CODE FRMLOCATION" & _
                    '        "               FROM GATE_OUTWARD_REG_HDR HDR" & _
                    '        "               WHERE HDR.UNIT_CODE='" + gstrUNITID + "' AND AGAINSTDOCUMENT='NRGP' and HDR.Doc_date > = '01 Apr 2022'   " & _
                    '        "        ) tmpTbl" & _
                    '        "        where CONVERT(VARCHAR(20),RH.DOC_NO)=tmpTbl.DOCUMENT_NO AND tmpTbl.FRMLOCATION=RH.From_Location" & _
                    '        "        AND tmpTbl.DOCUMENT_TYPE ='NRGP' " & _
                    '        " )" & _
                    '        " group by RH.Doc_No, RH.From_Location, RH.UNIT_CODE, RH.RGP_Date,VM.Vendor_Code, CM.Customer_Code, RH.Location_To_Type, VM.Vendor_name, CM.Cust_Name"

                    strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Document(s)", 0, 1, strSelectedRec)
                    If Not IsNothing(strHelp) AndAlso strHelp.Length < 8 Then
                        MessageBox.Show("No record found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return
                    ElseIf Not IsNothing(strHelp) AndAlso strHelp.Length = 7 Then
                        MessageBox.Show("Select atleast one document.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return
                    End If
                    If Not IsNothing(strHelp) AndAlso strHelp.Length = 8 Then
                        .Row = rNo
                        .Col = Enum_DocSel.Col_DocType
                        strSQL = "SELECT '" + Convert.ToString(.CellTag) + "' RowNo, '" + Convert.ToString(.Text) + "' InvoiceType,RH.DOC_NO, CONVERT(VARCHAR(20),RH.RGP_DATE,106) DOC_DT, ISNULL(CM.CUSTOMER_CODE,'') CUSTOMERCODE" & _
                            " , ISNULL(VM.VENDOR_CODE,'') VENDORCODE, RH.FROM_LOCATION FromLocation, '' VendorWhCode, SUM(ISNULL(RD.ACTUAL_QUANTITY,0)) QUANTITY, 0 DocValue" & _
                            " FROM RGP_HDR RH " & _
                            " inner join RGP_Dtl RD on RH.UNIT_CODE=RD.UNIT_CODE and RH.Doc_No=RD.Doc_No and RD.From_Location=RH.From_Location and RH.Doc_Type=RD.Doc_Type" & _
                            " LEFT JOIN VENDOR_MST VM ON RH.TO_LOCATION=VM.VENDOR_LOC AND RH.UNIT_CODE=VM.UNIT_CODE AND RH.LOCATION_TO_TYPE='V'" & _
                            " LEFT JOIN CUSTOMER_MST CM ON RH.TO_LOCATION=CM.CUST_LOCATION AND RH.UNIT_CODE=CM.UNIT_CODE AND RH.LOCATION_TO_TYPE='C'" & _
                            " WHERE RH.UNIT_CODE='" + gstrUNITID + "' AND RH.NRGP_CANCELLED=0 AND ISNULL(RH.AUTHORIZED_CODE,'')<>'' AND RH.LOCATION_TO_TYPE IN ('V','C')" & _
                            " AND RH.DOC_NO IN ('" + strHelp(strHelp.Length - 1).Replace("|", "','") + "')" & _
                            " AND RD.ACTUAL_QUANTITY>0 AND RH.RGP_DATE BETWEEN '" + dtpTripDt.Value.AddDays(-1 * noOfDays).ToString("dd MMM yyyy") + "' AND '" + dtpTripDt.Value.ToString("dd MMM yyyy") + "'" & _
                            " AND RH.DOC_TYPE=23" & _
                            " GROUP BY RH.DOC_NO, RH.FROM_LOCATION, RH.UNIT_CODE, RH.RGP_DATE,VM.VENDOR_CODE, CM.CUSTOMER_CODE, RH.LOCATION_TO_TYPE, VM.VENDOR_NAME, CM.CUST_NAME"
                        Dim dt As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                            For Each row As DataRow In dtDocDtl.DefaultView.ToTable(True, "InvoiceType", "RowNo", "Doc_No", "FromLocation").Select("RowNo in ('" + Convert.ToString(.CellTag) + "')")
                                If dt.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'").Length > 0 Then
                                    dt.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'")(0).Delete()
                                Else
                                    For Each r As DataRow In dtDocDtl.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'")
                                        r.Delete()
                                    Next
                                End If
                            Next
                            dt.AcceptChanges()
                            dtDocDtl.AcceptChanges()
                            Dim insertRow As DataRow
                            errMsg = String.Empty
                            For Each row As DataRow In dt.Rows
                                If Not String.IsNullOrEmpty(row("CUSTOMERCODE").ToString().Trim()) And Not isContractExist(row("CUSTOMERCODE").ToString().Trim(), txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                    If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                        errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                    End If
                                    Continue For
                                End If
                                If rbTransporterVeh.Checked Then
                                    If RadioButtonMonthly.Checked Then
                                        If Not isContractExistMonthly(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                    End If
                                    If RadioButtonCourier.Checked Then
                                        If Not isContractExistSize(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                        If Not isContractExistWeight(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                    End If
                                End If
                                insertRow = dtDocDtl.NewRow
                                insertRow("RowNo") = row("RowNo")
                                insertRow("InvoiceType") = row("InvoiceType")
                                insertRow("Doc_No") = row("DOC_NO")
                                insertRow("Doc_Dt") = row("DOC_DT")
                                insertRow("CustomerCode") = row("CUSTOMERCODE")
                                insertRow("VendorCode") = row("VENDORCODE")
                                insertRow("FromLocation") = row("FromLocation")
                                insertRow("VendorWhCode") = row("VendorWhCode")
                                insertRow("Quantity") = row("QUANTITY")
                                insertRow("DocValue") = row("DOCVALUE")
                                dtDocDtl.Rows.Add(insertRow)
                            Next
                            If dtDocDtl.Rows.Count > 0 Then
                                fillGrid()
                            End If
                            If Not String.IsNullOrEmpty(errMsg) Then
                                If rbTransporterVeh.Checked Then
                                    If RadioButtonCourier.Checked Then
                                        errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + "." + errMsg
                                    Else
                                        errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + " and Vehicle Category-" + txtVehicleCategory.Text.Trim() + "." + errMsg
                                    End If
                                Else
                                    errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + " and Vehicle Category-" + txtVehicleCategory.Text.Trim() + "." + errMsg
                                End If
                                MessageBox.Show(errMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            End If

                            .Row = rNo
                            .Col = Enum_DocSel.Col_DocNo
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    End If
                ElseIf .Text = "RGP" Then
                    fillRGPDocInTmp()

                    strSQL = "SELECT DOC_NO, FROM_LOCATION,CONVERT(VARCHAR(20), DOC_DT, 106) DOC_DT, INVOICE_TYPE," & strEwayBillType & " CUSTOMERCODE, VENDORCODE, CUST_NAME as [NAME], QUANTITY as TOTALQUANTITY   " & _
        " FROM TMP_FREIGHT_TRIP_SELECT_DOC WHERE UNIT_CODE='" + gstrUNITID + "' AND IP_ADDRESS='" + gstrIpaddressWinSck + "' and DOCTYPE = 'RGP'  "

                    strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Document(s)", 1, 1, strSelectedRec)
                    If Not IsNothing(strHelp) AndAlso strHelp.Length < 8 Then
                        MessageBox.Show("No record found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return
                    ElseIf Not IsNothing(strHelp) AndAlso strHelp.Length = 8 Then
                        MessageBox.Show("Select atleast one document.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return
                    End If
                    If Not IsNothing(strHelp) AndAlso strHelp.Length = 9 Then
                        .Row = rNo
                        .Col = Enum_DocSel.Col_DocType
                        strSQL = "SELECT '" + Convert.ToString(.CellTag) + "' RowNo, '" + Convert.ToString(.Text) + "' InvoiceType,RH.DOC_NO, CONVERT(VARCHAR(20),RH.RGP_DATE,106) DOC_DT, ISNULL(CM.CUSTOMER_CODE,'') CUSTOMERCODE" & _
                            " , ISNULL(VM.VENDOR_CODE,'') VENDORCODE, RH.FROM_LOCATION FromLocation, '' VendorWhCode, SUM(ISNULL(RD.ACTUAL_QUANTITY,0)) QUANTITY, 0 DocValue" & _
                            " FROM RGP_HDR RH " & _
                            " inner join RGP_Dtl RD on RH.UNIT_CODE=RD.UNIT_CODE and RH.Doc_No=RD.Doc_No and RD.From_Location=RH.From_Location and RH.Doc_Type=RD.Doc_Type" & _
                            " LEFT JOIN VENDOR_MST VM ON RH.TO_LOCATION=VM.VENDOR_LOC AND RH.UNIT_CODE=VM.UNIT_CODE AND RH.LOCATION_TO_TYPE='V'" & _
                            " LEFT JOIN CUSTOMER_MST CM ON RH.TO_LOCATION=CM.CUST_LOCATION AND RH.UNIT_CODE=CM.UNIT_CODE AND RH.LOCATION_TO_TYPE='C'" & _
                            " WHERE RH.UNIT_CODE='" + gstrUNITID + "' AND RH.NRGP_CANCELLED=0 AND ISNULL(RH.AUTHORIZED_CODE,'')<>'' AND RH.LOCATION_TO_TYPE IN ('V','C')" & _
                            " AND RH.DOC_NO IN ('" + strHelp(strHelp.Length - 1).Replace("|", "','") + "')" & _
                            " AND RD.ACTUAL_QUANTITY>0 AND RH.RGP_DATE BETWEEN '" + dtpTripDt.Value.AddDays(-1 * noOfDays).ToString("dd MMM yyyy") + "' AND '" + dtpTripDt.Value.ToString("dd MMM yyyy") + "'" & _
                            " AND RH.DOC_TYPE=22" & _
                            " GROUP BY RH.DOC_NO, RH.FROM_LOCATION, RH.UNIT_CODE, RH.RGP_DATE,VM.VENDOR_CODE, CM.CUSTOMER_CODE, RH.LOCATION_TO_TYPE, VM.VENDOR_NAME, CM.CUST_NAME"
                        Dim dt As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                            For Each row As DataRow In dtDocDtl.DefaultView.ToTable(True, "InvoiceType", "RowNo", "Doc_No", "FromLocation").Select("RowNo in ('" + Convert.ToString(.CellTag) + "')")
                                If dt.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'").Length > 0 Then
                                    dt.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'")(0).Delete()
                                Else
                                    For Each r As DataRow In dtDocDtl.Select("RowNo in ('" + Convert.ToString(row("RowNo")) + "') and Doc_No='" + Convert.ToString(row("Doc_No")) + "' and FromLocation='" + Convert.ToString(row("FromLocation")) + "'")
                                        r.Delete()
                                    Next
                                End If
                            Next
                            dt.AcceptChanges()
                            dtDocDtl.AcceptChanges()
                            Dim insertRow As DataRow
                            errMsg = String.Empty
                            For Each row As DataRow In dt.Rows
                                If Not String.IsNullOrEmpty(row("CUSTOMERCODE").ToString().Trim()) And Not isContractExist(row("CUSTOMERCODE").ToString().Trim(), txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                    If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                        errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                    End If
                                    Continue For
                                End If
                                If rbTransporterVeh.Checked Then
                                    If RadioButtonMonthly.Checked Then
                                        If Not isContractExistMonthly(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                    End If
                                    If RadioButtonCourier.Checked Then
                                        If Not isContractExistSize(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                        If Not isContractExistWeight(txtTransporterCode.Text.Trim(), dtpTripDt.Value.ToString("dd MMM yyyy")) Then
                                            If Not errMsg.Contains(row("CUSTOMERCODE").ToString().Trim()) Then
                                                errMsg = errMsg + vbCrLf + row("CUSTOMERCODE").ToString().Trim()
                                            End If
                                            Continue For
                                        End If
                                    End If
                                End If
                                insertRow = dtDocDtl.NewRow
                                insertRow("RowNo") = row("RowNo")
                                insertRow("InvoiceType") = row("InvoiceType")
                                insertRow("Doc_No") = row("DOC_NO")
                                insertRow("Doc_Dt") = row("DOC_DT")
                                insertRow("CustomerCode") = row("CUSTOMERCODE")
                                insertRow("VendorCode") = row("VENDORCODE")
                                insertRow("FromLocation") = row("FromLocation")
                                insertRow("VendorWhCode") = row("VendorWhCode")
                                insertRow("Quantity") = row("QUANTITY")
                                insertRow("DocValue") = row("DOCVALUE")
                                dtDocDtl.Rows.Add(insertRow)
                            Next
                            If dtDocDtl.Rows.Count > 0 Then
                                fillGrid()
                            End If
                            If Not String.IsNullOrEmpty(errMsg) Then
                                If rbTransporterVeh.Checked Then
                                    If RadioButtonCourier.Checked Then
                                        errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + "." + errMsg
                                    Else
                                        errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + " and Vehicle Category-" + txtVehicleCategory.Text.Trim() + "." + errMsg
                                    End If
                                Else
                                    errMsg = "Contract not defined for below selected customer(s) against transporter-" + txtTransporterCode.Text.Trim() + " and Vehicle Category-" + txtVehicleCategory.Text.Trim() + "." + errMsg
                                End If
                                MessageBox.Show(errMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            End If
                            .Row = rNo
                            .Col = Enum_DocSel.Col_DocNo
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fillInvoiceDocInTmp()
        Try
            Dim cmd As SqlCommand = New SqlCommand()
            cmd.Connection = SqlConnectionclass.GetConnection()
            cmd.CommandText = "PROC_FREIGHT_TRIP_GET_DOCUMENTS"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 0
            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
            cmd.Parameters.AddWithValue("@DOC_TYPE", "INVOICE")
            cmd.Parameters.AddWithValue("@TRAN_TYPE", "GET_INVOICES")
            cmd.Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
            cmd.Parameters.AddWithValue("@TRIP_DT", dtpTripDt.Value.ToString("dd MMM yyyy"))
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fillRGPDocInTmp(Optional ByVal DOCTYPE As String = "RGP")
        Try
            Dim cmd As SqlCommand = New SqlCommand()
            cmd.Connection = SqlConnectionclass.GetConnection()
            cmd.CommandText = "PROC_FREIGHT_TRIP_GET_DOCUMENTS"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 0
            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
            cmd.Parameters.AddWithValue("@DOC_TYPE", DOCTYPE)
            cmd.Parameters.AddWithValue("@TRAN_TYPE", "GET_" & DOCTYPE)
            cmd.Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
            cmd.Parameters.AddWithValue("@TRIP_DT", dtpTripDt.Value.ToString("dd MMM yyyy"))
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' check if contract exists for customer or vendor ware house for selected transporter in defined date.
    ''' </summary>
    ''' <param name="strCustWhCode">customer code and vendor ware house code</param>
    ''' <param name="strTransporterCode">Transporter code</param>
    ''' <param name="effDate">Trip date</param>
    ''' <returns>true if record exists else returns false.</returns>
    Private Function isContractExist(ByVal strCustWhCode As String, ByVal strTransporterCode As String, ByVal effDate As String) As Boolean
        Try
            If rbTransporterVeh.Checked Then
                If RadioButtonTrip.Checked Then
                    Return IsRecordExists("SELECT * FROM FREIGHT_CONTRACT_DTL FCD " & _
                                    " INNER JOIN FREIGHT_CONTRACT_HDR FCH ON FCD.CONTRACT_ID=FCH.CONTRACT_ID AND FCD.UNIT_CODE=FCH.UNIT_CODE" & _
                                    " AND FCH.TRANSPORTER_CODE='" + strTransporterCode + "' AND FCH.AUTHORIZED_STATUS='AUTHORIZED' AND FCH.CONTRACT_TYPE='" + _contractTypeTripBased + "' " & _
                                    " WHERE FCD.CUST_VEND_WH_CODE='" + strCustWhCode + "' AND FCD.UNIT_CODE='" + gstrUNITID + "' AND '" + effDate + "' BETWEEN FCD.EFFECTIVE_FROM AND FCD.EFFECTIVE_TO " & _
                                    " AND FCD.VEHICLE_CATEGORY_CODE='" + txtVehicleCategory.Text.Trim() + "' AND FCH.CONTRACT_TYPE='" + _contractTypeTripBased + "' AND ISNULL(FCD.CONTRACT_STATUS,'')<>'AUTHORIZED' ")
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Function isContractExistMonthly(ByVal strTransporterCode As String, ByVal effDate As String) As Boolean
        Try
            If rbTransporterVeh.Checked Then
                If RadioButtonMonthly.Checked Then
                    Return IsRecordExists("SELECT * FROM FREIGHT_CONTRACT_DTL_MONTHLY FCD " & _
                                    " INNER JOIN FREIGHT_CONTRACT_HDR FCH ON FCD.CONTRACT_ID=FCH.CONTRACT_ID AND FCD.UNIT_CODE=FCH.UNIT_CODE" & _
                                    " AND FCH.TRANSPORTER_CODE='" + strTransporterCode + "' AND FCH.AUTHORIZED_STATUS='AUTHORIZED' AND FCH.CONTRACT_TYPE='" + _contractTypeMonthlyBased + "' " & _
                                    " WHERE FCD.UNIT_CODE='" + gstrUNITID + "' AND '" + effDate + "' BETWEEN FCD.EFFECTIVE_FROM AND FCD.EFFECTIVE_TO " & _
                                    " AND FCD.VEHICLE_CATEGORY_CODE='" + txtVehicleCategory.Text.Trim() + "' AND FCH.CONTRACT_TYPE='" + _contractTypeMonthlyBased + "' AND ISNULL(FCD.CONTRACT_STATUS,'')<>'AUTHORIZED' ")
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Function isContractExistSize(ByVal strTransporterCode As String, ByVal effDate As String) As Boolean
        Try
            If rbTransporterVeh.Checked Then
                If RadioButtonCourier.Checked Then
                    Return IsRecordExists("SELECT * FROM FREIGHT_CONTRACT_DTL_SIZE_WEIGHT FCD " & _
                                    " INNER JOIN FREIGHT_CONTRACT_HDR FCH ON FCD.CONTRACT_ID=FCH.CONTRACT_ID AND FCD.UNIT_CODE=FCH.UNIT_CODE" & _
                                    " AND FCH.TRANSPORTER_CODE='" + strTransporterCode + "' AND FCH.AUTHORIZED_STATUS='AUTHORIZED' AND FCH.CONTRACT_TYPE='" + _contractTypeCourierBased + "' " & _
                                    " WHERE FCD.UNIT_CODE='" + gstrUNITID + "' AND '" + effDate + "' BETWEEN FCD.EFFECTIVE_FROM AND FCD.EFFECTIVE_TO " & _
                                    " AND FCH.CONTRACT_TYPE='" + _contractTypeCourierBased + "' AND ISNULL(FCD.CONTRACT_STATUS,'')<>'AUTHORIZED' AND FCD.TRANSPORT_MODE='" + TextBoxTransportMode.Text.Trim() + "' AND FCD.ZONE='" + TextBoxZone.Text.Trim() + "' AND FCD.ENTRY_TYPE='S' AND " & Val(TextBoxSize.Text) & " Between FCD.SIZE_CUBIC_FEET_FROM AND FCD.SIZE_CUBIC_FEET_TO")
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Function isContractExistWeight(ByVal strTransporterCode As String, ByVal effDate As String) As Boolean
        Try
            If rbTransporterVeh.Checked Then
                If RadioButtonCourier.Checked Then
                    Return IsRecordExists("SELECT * FROM FREIGHT_CONTRACT_DTL_SIZE_WEIGHT FCD " & _
                                    " INNER JOIN FREIGHT_CONTRACT_HDR FCH ON FCD.CONTRACT_ID=FCH.CONTRACT_ID AND FCD.UNIT_CODE=FCH.UNIT_CODE" & _
                                    " AND FCH.TRANSPORTER_CODE='" + strTransporterCode + "' AND FCH.AUTHORIZED_STATUS='AUTHORIZED' AND FCH.CONTRACT_TYPE='" + _contractTypeCourierBased + "' " & _
                                    " WHERE FCD.UNIT_CODE='" + gstrUNITID + "' AND '" + effDate + "' BETWEEN FCD.EFFECTIVE_FROM AND FCD.EFFECTIVE_TO " & _
                                    " AND FCH.CONTRACT_TYPE='" + _contractTypeCourierBased + "' AND ISNULL(FCD.CONTRACT_STATUS,'')<>'AUTHORIZED' AND FCD.TRANSPORT_MODE='" + TextBoxTransportMode.Text.Trim() + "' AND FCD.ZONE='" + TextBoxZone.Text.Trim() + "' AND FCD.ENTRY_TYPE='W' AND " & Val(ctlWeight.Text) & " Between FCD.WEIGHT_KG_FROM AND FCD.WEIGHT_KG_TO")
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Sub rbVeh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTransporterVeh.CheckedChanged, rbCustVeh.CheckedChanged
        Try
            If rbTransporterVeh.Checked Then
                txtVehType.Text = String.Empty
                txtVehType.Enabled = False
                txtCarrName.Text = String.Empty
                txtCarrName.Enabled = False
                txtTransporterCode.Text = String.Empty
                txtTransporterCode.Enabled = True
                txtVehicleCategory.Text = String.Empty
                txtVehicleCategory.Enabled = True
                BtnHelpTransporter.Enabled = True
                BtnHelpVehCategory.Enabled = True
                RadioButtonTrip.Enabled = True
                RadioButtonMonthly.Enabled = True
                RadioButtonCourier.Enabled = True
                RadioButtonTrip.Checked = True
                TextBoxSize.Text = "0.00"
                ctlWeight.Text = "0.00"
                ctlWeight.Enabled = False
                TextBoxLenghtFT.Enabled = False
                TextBoxWidthFT.Enabled = False
                TextBoxHeightFT.Enabled = False
                TextBoxLengthInch.Enabled = False
                TextBoxWidthInch.Enabled = False
                TextBoxHeightInch.Enabled = False
                TextBoxZone.Enabled = False
                TextBoxLenghtFT.Text = ""
                TextBoxWidthFT.Text = ""
                TextBoxHeightFT.Text = ""
                TextBoxLengthInch.Text = ""
                TextBoxWidthInch.Text = ""
                TextBoxHeightInch.Text = ""
                TextBoxZone.Text = ""
                ButtonZone.Enabled = False
                TextBoxTransportMode.Enabled = False
                TextBoxTransportMode.Text = ""
                ButtonTransportMode.Enabled = False
                clearGrid()
                dtpTripDt.Focus()
                dtpTripDt.Focus()
                txtVehicleNo.Enabled = False
                'added by priti on 16 march 2020 to add vehicle help box
                If mblnAllowTransporterfromMaster Then
                    'txtVehicleNo.Enabled = True
                    cmdVehicleCodeHelp.Visible = True
                    cmdVehicleCodeHelp.Enabled = True
                    txtVehicleCategory.Enabled = False
                    BtnHelpVehCategory.Enabled = False
                    txtVehicleNo.Text = ""
                Else
                    txtVehicleNo.Enabled = True
                    cmdVehicleCodeHelp.Visible = False
                    cmdVehicleCodeHelp.Enabled = False
                    txtVehicleCategory.Enabled = True
                    BtnHelpVehCategory.Enabled = True
                End If

            Else
                txtTransporterCode.Text = String.Empty
                txtTransporterCode.Enabled = False
                txtVehicleCategory.Text = String.Empty
                txtVehicleCategory.Enabled = False
                BtnHelpTransporter.Enabled = False
                BtnHelpVehCategory.Enabled = False
                txtVehType.Text = String.Empty
                txtVehType.Enabled = True
                txtCarrName.Text = String.Empty
                txtCarrName.Enabled = True
                RadioButtonTrip.Enabled = False
                RadioButtonMonthly.Enabled = False
                RadioButtonCourier.Enabled = False
                RadioButtonTrip.Checked = False
                RadioButtonMonthly.Checked = False
                RadioButtonCourier.Checked = False
                TextBoxSize.Text = "0.00"
                ctlWeight.Text = "0.00"
                ctlWeight.Enabled = False
                TextBoxLenghtFT.Enabled = False
                TextBoxWidthFT.Enabled = False
                TextBoxHeightFT.Enabled = False
                TextBoxLengthInch.Enabled = False
                TextBoxWidthInch.Enabled = False
                TextBoxHeightInch.Enabled = False
                TextBoxZone.Enabled = False
                TextBoxLenghtFT.Text = ""
                TextBoxWidthFT.Text = ""
                TextBoxHeightFT.Text = ""
                TextBoxLengthInch.Text = ""
                TextBoxWidthInch.Text = ""
                TextBoxHeightInch.Text = ""
                TextBoxZone.Text = ""
                ButtonZone.Enabled = False
                TextBoxTransportMode.Enabled = False
                TextBoxTransportMode.Text = ""
                ButtonTransportMode.Enabled = False
                clearGrid()
                AddNewRow()
                dtpTripDt.Focus()
                dtpTripDt.Focus()
                'added by priti on 16 march 2020 to add vehicle help box
                txtVehicleNo.Enabled = True
                If mblnAllowTransporterfromMaster Then
                    'txtVehicleNo.Enabled = True
                    cmdVehicleCodeHelp.Visible = False
                    cmdVehicleCodeHelp.Enabled = False
                    txtVehicleCategory.Enabled = False
                    BtnHelpVehCategory.Enabled = False
                    txtVehicleNo.Text = ""
                End If

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub setAdHoc()
        Try
            Dim strSQL As String
            strSQL = "Select isAdhocTripActive from POConfig_Mst where unit_code = '" + gstrUNITID + "'"

            blnIsAdhoc = SqlConnectionclass.ExecuteScalar(strSQL)
            If blnIsAdhoc = True Then
                rdoAdHoc.Visible = True
                grpAdhocContractType.Visible = False
                grpReqNo.Visible = False
                btnAdhocDetails.Visible = False
                BtnViewdoc.Visible = False
            Else
                rdoAdHoc.Visible = False
                grpAdhocContractType.Visible = False
                grpReqNo.Visible = False
                btnAdhocDetails.Visible = False
                BtnViewdoc.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub

#End Region

    Private Sub RadioButtonTrip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonTrip.CheckedChanged, RadioButtonMonthly.CheckedChanged, RadioButtonCourier.CheckedChanged
        If RadioButtonCourier.Checked Then
            txtVehicleCategory.Text = String.Empty
            txtVehicleCategory.Enabled = False
            txtTransporterCode.Text = String.Empty
            txtTransporterName.Text = String.Empty
            BtnHelpVehCategory.Enabled = False
            TextBoxSize.Text = "0.00"
            ctlWeight.Text = "0.00"
            ctlWeight.Enabled = True
            TextBoxLenghtFT.Enabled = True
            TextBoxWidthFT.Enabled = True
            TextBoxHeightFT.Enabled = True
            TextBoxLengthInch.Enabled = True
            TextBoxWidthInch.Enabled = True
            TextBoxHeightInch.Enabled = True
            TextBoxZone.Enabled = True
            TextBoxLenghtFT.Text = ""
            TextBoxWidthFT.Text = ""
            TextBoxHeightFT.Text = ""
            TextBoxLengthInch.Text = ""
            TextBoxWidthInch.Text = ""
            TextBoxHeightInch.Text = ""
            TextBoxZone.Text = ""
            ButtonZone.Enabled = True
            TextBoxTransportMode.Enabled = True
            TextBoxTransportMode.Text = ""
            ButtonTransportMode.Enabled = True

        Else
            txtVehicleCategory.Text = String.Empty
            txtVehicleCategory.Enabled = True
            txtTransporterCode.Text = String.Empty
            txtTransporterName.Text = String.Empty
            BtnHelpVehCategory.Enabled = True
            TextBoxSize.Text = "0.00"
            ctlWeight.Text = "0.00"
            ctlWeight.Enabled = False
            TextBoxLenghtFT.Enabled = False
            TextBoxWidthFT.Enabled = False
            TextBoxHeightFT.Enabled = False
            TextBoxLengthInch.Enabled = False
            TextBoxWidthInch.Enabled = False
            TextBoxHeightInch.Enabled = False
            TextBoxZone.Enabled = False
            TextBoxLenghtFT.Text = ""
            TextBoxWidthFT.Text = ""
            TextBoxHeightFT.Text = ""
            TextBoxLengthInch.Text = ""
            TextBoxWidthInch.Text = ""
            TextBoxHeightInch.Text = ""
            TextBoxZone.Text = ""
            ButtonZone.Enabled = False
            TextBoxTransportMode.Enabled = False
            TextBoxTransportMode.Text = ""
            ButtonTransportMode.Enabled = False
            dtpTripDt.Focus()

        End If
    End Sub

    Private Sub ButtonZone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonZone.Click
        Dim strSQL As String = String.Empty
        Dim strHelp As String()
        Try
            If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strSQL = "SELECT DISTINCT ZONE AS CODE,ZONE AS DESCRIPTION FROM STATE_MST WHERE ISNULL(ZONE,'')<>'' ORDER BY ZONE"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Zone", 1, 0, TextBoxZone.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 AndAlso Not IsNothing(strHelp(1)) Then
                        TextBoxZone.Text = strHelp(0).Trim
                        AddNewRow()
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxZone_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxZone.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                ButtonZone_Click(ButtonZone, New EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxZone_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxZone.KeyPress
        Try
            If e.KeyChar = "'" Then
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxZone_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBoxZone.Validating
        Dim strSQL As String = String.Empty
        Try
            If Len(TextBoxZone.Text.Trim) > 0 Then
                strSQL = "SELECT 1 FROM STATE_MST WHERE ISNULL(ZONE,'')<>'' And ZONE='" & TextBoxZone.Text.Trim & "'"
                If IsRecordExists(strSQL) Then
                    TextBoxZone.Text = TextBoxZone.Text.Trim.ToUpper
                Else
                    MsgBox("Please enter valid Zone!", MsgBoxStyle.Information, "eMPro")
                    TextBoxZone.Text = ""
                    TextBoxZone.Focus()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxLenghtFT_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxLenghtFT.KeyPress, TextBoxWidthFT.KeyPress, TextBoxHeightFT.KeyPress
        Try
            If Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxLengthInch_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxLengthInch.KeyPress, TextBoxWidthInch.KeyPress, TextBoxHeightInch.KeyPress
        Try
            If Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxLenghtFT_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxLenghtFT.TextChanged, TextBoxWidthFT.TextChanged, TextBoxHeightFT.TextChanged, TextBoxLengthInch.TextChanged, TextBoxWidthInch.TextChanged, TextBoxHeightInch.TextChanged
        Try
            Dim lengthInInch As Double = 0
            Dim widthInInch As Double = 0
            Dim heightInInch As Double = 0
            lengthInInch = (Val(TextBoxLenghtFT.Text) * 12) + Val(TextBoxLengthInch.Text)
            widthInInch = (Val(TextBoxWidthFT.Text) * 12) + Val(TextBoxWidthInch.Text)
            heightInInch = (Val(TextBoxHeightFT.Text) * 12) + Val(TextBoxHeightInch.Text)
            TextBoxSize.Text = ((lengthInInch * widthInInch * heightInInch) / 1728).ToString("0.00")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxLengthInch_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBoxLengthInch.Validating, TextBoxWidthInch.Validating, TextBoxHeightInch.Validating
        Try
            Dim txtBox As New TextBox
            txtBox = DirectCast(sender, TextBox)
            If Val(txtBox.Text) > 11 Then
                MsgBox("Enter Value should not be greater than 11!", MsgBoxStyle.Information, "eMPro")
                txtBox.Text = ""
                txtBox.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ButtonTransportMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonTransportMode.Click
        Dim strSQL As String = String.Empty
        Dim strHelp As String()
        Try
            If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strSQL = "SELECT DISTINCT DESCR As CODE,DESCR As DESCRIPTION FROM LISTS WHERE KEY1='FREIGHT' AND KEY2='TRANSPORT_MODE' AND UNIT_CODE='" & gstrUNITID & "'"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Transport Mode", 1, 0, TextBoxTransportMode.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 AndAlso Not IsNothing(strHelp(1)) Then
                        TextBoxTransportMode.Text = strHelp(0).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxTransportMode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxTransportMode.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                ButtonTransportMode_Click(ButtonTransportMode, New EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxTransportMode_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxTransportMode.KeyPress
        Try
            If e.KeyChar = "'" Then
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TextBoxTransportMode_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBoxTransportMode.Validating
        Dim strSQL As String = String.Empty
        Try
            If Len(TextBoxTransportMode.Text.Trim) > 0 Then
                strSQL = "SELECT 1 FROM LISTS WHERE KEY1='FREIGHT' AND KEY2='TRANSPORT_MODE' AND UNIT_CODE='" & gstrUNITID & "' AND DESCR='" & TextBoxTransportMode.Text.Trim() & "'"
                If IsRecordExists(strSQL) Then
                    TextBoxTransportMode.Text = TextBoxTransportMode.Text.Trim.ToUpper
                Else
                    MsgBox("Please enter valid Transport Mode!", MsgBoxStyle.Information, "eMPro")
                    TextBoxTransportMode.Text = ""
                    TextBoxTransportMode.Focus()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdVehicleCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVehicleCodeHelp.Click
        'select transporter_code as [Transporter Code],Transporter_name as [Transporter Name],Vehicle_Type as [Vehicle Type],vehicle_no as [Vehicle No] from vehicle_mst where active=1
        Dim strSql As String = ""
        Dim strVehicle As String = ""
        Dim varRetVal As Object
        On Error GoTo ErrHandler
        With txtVehicleNo
            If Len(.Text) = 0 Then

                'varRetVal = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "  and Group_Customer=1 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")

                varRetVal = ShowList(1, .MaxLength, "", "vehicle_no", "Transporter_name", "vehicle_mst", " and active=1  and transporter_code='" & txtTransporterCode.Text & "'", "Help", "", 1, 0, "transporter_code")
                If varRetVal = "-1" Then
                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    .Text = ""
                Else
                    .Text = varRetVal
                    strSql = "SELECT Top 1 vehicle_type FROM vehicle_mst WHERE UNIT_CODE= '" & gstrUNITID & "' AND transporter_code = '" & txtTransporterCode.Text.Trim & "'  and vehicle_no='" & txtVehicleNo.Text & "'"
                    txtVehicleCategory.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))
                    '.Text = strVehicle
                End If
            Else
                varRetVal = ShowList(1, .MaxLength, , "vehicle_no", "Transporter_name", "vehicle_mst", " and active=1 and transporter_code='" & txtTransporterCode.Text & "'", "Help", "", 1, 0, "transporter_code")
                If varRetVal = "-1" Then
                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    .Text = ""
                Else
                    .Text = varRetVal
                    strSql = "SELECT Top 1 vehicle_type FROM vehicle_mst WHERE UNIT_CODE= '" & gstrUNITID & "' AND transporter_code = '" & txtTransporterCode.Text.Trim & "'  and vehicle_no='" & txtVehicleNo.Text & "'"
                    txtVehicleCategory.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))
                    '.Text = strVehicle
                End If
            End If

            .Focus()
            clearGrid()
            AddNewRow()
        End With

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub rdoAdHoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoAdHoc.CheckedChanged
        Try
            grpAdhocContractType.Enabled = False
            rdoInward.Checked = True
            txtReqNo.Text = ""
            If rdoAdHoc.Checked = True Then
                grpAdhocContractType.Visible = True
                grpReqNo.Visible = True
                btnAdhocDetails.Visible = True
                BtnViewdoc.Visible = True
                btnAdhocDetails.Enabled = True
                BtnViewdoc.Enabled = True
            Else
                grpAdhocContractType.Visible = False
                grpReqNo.Visible = False
                btnAdhocDetails.Visible = False
                BtnViewdoc.Visible = False
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnHelpReqNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelpReqNo.Click

        Dim strSQL As String = String.Empty
        Dim strHelp As String()

        Try
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With

            strSQL = "Declare @ExpDays int =900000 " & vbCrLf
            strSQL = strSQL & "Select @ExpDays = NoOfDays " & vbCrLf
            strSQL = strSQL & "From Lists " & vbCrLf
            strSQL = strSQL & "Where Key1 = 'Ad_Hoc_Request_Life_Days' And 1 = (Case When Unit_code ='GLOBAL' Then 1 Else  Case When UNIT_CODE = '" & gstrUNITID & "' Then 1 Else 0 End End) " & vbCrLf
            strSQL = strSQL & "select requestNo, TypeOfShipment, ReqDate, TransporterID,VehicleType from vw_getAdhocReqNoForTripGeneration where unitcode = '" + gstrUNITID + "' And DateDiff(Day,ReqDate,GETDATE()) <= @ExpDays "
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Adhoc Request No", 1, 0, txtReqNo.Text.Trim)
            If Not IsNothing(strHelp) Then
                If strHelp.Length >= 3 Then
                    txtReqNo.Text = strHelp(0).Trim

                    If strHelp(1) = "I" Then
                        rdoInward.Checked = True
                    ElseIf strHelp(1) = "O" Then
                        rdoOutward.Checked = True
                    End If

                    txtTransporterCode.Text = strHelp(3)
                    ' adhoc phase 2
                    txtVehicleCategory.Text = strHelp(4)
                    BtnHelpVehCategory.Enabled = False
                    txtVehicleCategory.Enabled = False
                Else
                    MsgBox("No Request No Found.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                End If
            Else
                MsgBox("No Request No Found.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, ResolveResString(100))
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Sub
            End If


            strSQL = " SELECT VM.VENDOR_NAME FROM VENDOR_MST VM WHERE VM.UNIT_CODE='" + gstrUNITID + "'" & _
                       " AND VM.VENDOR_CODE = '" + txtTransporterCode.Text + "'"
            txtTransporterName.Text = SqlConnectionclass.ExecuteScalar(strSQL)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnAdhocDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdhocDetails.Click
        Try
            Dim frmObj As frmAdHocRequestDetails
            Dim strSQL As String = ""
            Dim oRDR As SqlDataReader

            strSQL = "select Quoted_Price, NegotiatedPrice, ChargeableWeight from vw_getAdhocReqNoForTripGeneration where unitcode = '" + gstrUNITID + "' and requestNo = '" + txtReqNo.Text + "'"
            oRDR = SqlConnectionclass.ExecuteReader(strSQL)

            frmObj = New frmAdHocRequestDetails

            If (oRDR.HasRows) Then
                oRDR.Read()
                frmObj.lblQuotedPrice.Text = oRDR("Quoted_Price").ToString()
                frmObj.lblApprovedPrice.Text = oRDR("NegotiatedPrice").ToString()
                frmObj.lblChargeableWt.Text = oRDR("ChargeableWeight").ToString()
                frmObj.lblTransporter.Text = txtTransporterCode.Text
            End If

            frmObj.lblRequestNo.Text = txtReqNo.Text
            frmObj.lblTransporter.Text = txtTransporterCode.Text

            frmObj.Show()
            frmObj.BringToFront()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub BtnViewdoc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnViewdoc.Click
        Dim strSql As String = String.Empty
        Try
            If txtReqNo.Text.Trim = "" Then
                MsgBox("Please select Request no.!", MsgBoxStyle.Exclamation, ResolveResString(100))
                Exit Sub
            End If
            Dim objfrmLGTTRN0002a As New frmLGTTRN0002a()
            objfrmLGTTRN0002a.GetSetReqNo = txtReqNo.Text
            objfrmLGTTRN0002a.GetSetUnitCode = gstrUNITID
            objfrmLGTTRN0002a.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Function ALLOWCUSTOMERSPECIFICREPORT_ROWSIZELARGE() As Boolean
        Dim RetValue As Boolean = False
        Try
            If ((Not String.IsNullOrEmpty(CustomerCode)) AndAlso CustomerCodes.All(Function(x) x.Equals(CustomerCode))) Then
                Dim result = SqlConnectionclass.ExecuteScalar("SELECT TOP 1 ALLOWCUSTOMERSPECIFICREPORT_ROWSIZELARGE FROM CUSTOMER_MST WHERE UNIT_CODE = '" + gstrUNITID + "' AND CUSTOMER_CODE ='" + CustomerCode + "' ")
                If result IsNot Nothing AndAlso result = True Then
                    RetValue = True
                End If
            End If
            Return RetValue
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return RetValue
        End Try

    End Function
End Class