Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Friend Class FRMMKTTRN0028HI
    Inherits System.Windows.Forms.Form
    '********************************************************************************************************
    'Copyright (c)  -   MIND
    'Name of module -   FRMMFGTRN0028.frm
    'Created By     -   Amal Doss.L
    'Created On     -   19-Feb-2007
    'description    -   Schedule Uploading as per DARWIN
    '               -   frmMKTTRN0028

    'Revision On    -   14-Mar-2007
    'Revised By     -   Shubhra Verma
    'Revised On     -   29-may-2007 and 30-may-2007
    'Revised By     -   Shubhra
    'Description    -   schedule uploading for COVISINT

    'Revised By         : Shubhra Verma
    'Revised On         : 07-06-2007
    'Issue ID           : 20085
    'Revision History   : USER CAN STRAIGHTLY SUPPLY PARTS/MATERIAL TO THE CONSIGNEE, NOT THROUGH WAREHOUSE

    'Revised By         : Shubhra Verma
    'Revised On         : 04-Aug-2007
    'Issue ID           : 20756
    'Revision History   : Whenever Uploading the Release File,It should check
    '                     Transmission number for the SenderID/Customer Code in the
    '                     Release File.If the last number is not equal to Current
    '                     (in new Release)-1 then Alert for the Same alongwith mail.
    '                     Also Uploading should be allowed after Authorisation for
    '                     uploading of the same.

    'Revised By         : Shubhra Verma
    'Revised On         : 01-JAN-2008
    'Issue ID           : 21968
    'Revision History   : USER CAN RESET THE CALLOFFNO SERIES

    'Revised By         : Shubhra Verma
    'Revised On         : 04 Mar 2008 to 07 Mar 2008
    'Issue id           : eMpro-20080306-13517
    'Revision History   :1 - There should be a provision of using daily pull
    '                    qty from release file parameter master as minimum
    '                    safety stock if daily pull qty check box is checked
    '                    in CDP form.
    '                    2 - Schedule Qty should be calculated on Bag Qty.

    'Revised By         : Shubhra Verma
    'Revised On         : 10 Mar 2008
    'Issue id           : eMpro-20080306-13571
    'Revision History   : System should upload file for multiple consignee
    '                     of same warehouse

    'Revised By         : Shubhra Verma
    'Revised On         : 18 Mar 2008 to 19 Mar 2008
    'Issue id           : eMpro-20080317-15057
    'Revision History   : System Should Use "Daily Pull Rate" from Warehouse
    '                     Stock Report as DailyPullQty, so that System can calculate
    '                     Consignee Wise Safety Stock.

    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 10 APR 2008 TO 18 APR 2008
    'ISSUE ID           : 'ISSUE ID - eMpro-20080410-17008

    '********************************************************************************************************
    'REVISED BY         : MANOJ VAISH
    'REVISED ON         : 16 JUL 2008
    'ISSUE ID           : 'ISSUE ID - eMpro-20080716-20353
    'REVISION HISTORY   : 'Some Issue Rectification reported by Mustafa

    '********************************************************************************************************
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 12 SEP 2008
    'ISSUE ID           : eMpro-20080911-21453
    '********************************************************************************************************
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 03 Mar 2009
    'ISSUE ID           : SouthUnits-20090302-28106 
    '                   : CDP output is not updated in the daily marketing schedule.
    '********************************************************************************************************
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 05 Mar 2009 to 09 Mar 2009
    'ISSUE ID           : eMpro-20090309-28458
    '                   : uploading of warehouse stock file and invoice file in txt format.
    '********************************************************************************************************
    'REVISED BY         : MANOJ VAISH
    'REVISED ON         : 04 JUN 2009
    'ISSUE ID           : eMpro-20090604-32080
    '                   : While uploading the schedule in dailymktschedule the revision no
    '                   : should be incremented.
    '********************************************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090611-32362
    'Revision Date      : 14 Jun 2009
    'History            : Add New field of RAN No for Edifact File -Hilex Nissan CSV File Genaration
    '****************************************************************************************
    '============================================================================================
    'Revised By         :   Shalini Singh
    'Revised On         :   23 Sep 2011
    'Reason             :   Ip address and C drive hard coded change
    'issue id           :   10140039
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 01 Nov 2012
    'ISSUE ID           : 10303249
    'Description        : if new schedule qty is less than despatch qty, then new sch qty = new sch qty + prev despatch qty
    '============================================================================================

    Dim mintFormIndex As Short
    Dim Obj_FSO As Scripting.FileSystemObject
    Dim Obj_EX As Excel.Application
    Dim Upload_FileType As String
    Dim Flag As Short
    Dim consignee As String
    Dim mblnDailymktUpdated, mblnfilemove As Boolean
    Dim Remarks As String
    Dim Bool_Not_File As Boolean = False
    Dim range As Excel.Range
    Dim bool_Validate_Cust As Boolean = False
    Dim bool_Validate_Cons As Boolean = False
    Public cust_code As String = ""
    Public file_name As String = ""
    Dim bln_dateCheck As Boolean = False
    Dim blnMsgforDate As Boolean = False
    Dim custFocus As Boolean = False
    Dim consFocus As Boolean = False
    Dim unitFocus As Boolean = False
    Dim mlngBAGQTY As Long
    Dim blnCOVISINTScheduleOverwrite As Boolean
    Dim mShipmentFlag As Boolean

    Private Enum Enum_Up
        Del_Dt = 1
        Cust_Drg_No
        Item_Code
        Item_Desc
        SftyStkPerDay
        SftyDays
        Sch_Qty
        Hidd_Dt
        wh_code
        CONSIGNEE_CODE
        RAN_No
    End Enum

    Private Enum enumWH
        Stock_dt = 1
        CustPartNo
        ItemCode
        Description
        AvlStk
        SaftyStk
        wh_code
    End Enum

    Private Function FN_Date_Conversion(ByRef Cell_Dt As String) As Object
        Dim T_Month, T_Date, T_Year As String
        Dim HOLIDAY As Short
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = String.Empty

        Try

            HOLIDAY = 1
            Cell_Dt = Replace(Cell_Dt, "'", "")
            If Len(Cell_Dt) >= 5 Then
                T_Date = Mid(Cell_Dt, Len(Cell_Dt) - 1, 2)
                T_Month = Mid(Cell_Dt, Len(Cell_Dt) - 3, 2)
                T_Year = Mid(Cell_Dt, 1, Len(Cell_Dt) - 4)
                If Len(T_Year) = 1 Then
                    T_Year = "200" & T_Year
                Else
                    T_Year = "20" & T_Year
                End If
                If T_Date = "00" Then T_Date = "01"
                If IsDate(T_Date & "/" & T_Month & "/" & T_Year) = True Then
                    FN_Date_Conversion = T_Date & "/" & T_Month & "/" & T_Year

                    If Mid(Cell_Dt, Len(Cell_Dt) - 1, 2) = "00" Then
                        mP_Connection.Execute("set dateformat 'dmy' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        strSql = "select work_flg from calendar_mkt_cust" & " where dt = convert(varchar,'" & FN_Date_Conversion & "',106) " & " and Cust_Code = '" & consignee & "' and UNIT_CODE='" & gstrUNITID & "'"
                        sqlRdr = SqlConnectionclass.ExecuteReader(strSql)

                        While sqlRdr("WORK_FLG") = True
                            T_Date = CStr(CDbl(T_Date) + 1)
                            FN_Date_Conversion = CDbl(T_Date) + 1 & "/" & T_Month & "/" & T_Year

                            'RSconsignee.ResultSetClose()
                            'RSconsignee = New ClsResultSetDB
                            'RSconsignee.GetResult("select work_flg from calendar_mkt_cust" & " where dt = convert(varchar,'" & FN_Date_Conversion & "',106) " & " and Cust_Code = '" & consignee & "' and UNIT_CODE='" & gstrUNITID & "'")
                        End While
                        'RSconsignee.ResultSetClose()

                    End If

                Else
                    FN_Date_Conversion = ""
                End If
            Else
                FN_Date_Conversion = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Function

    Private Function FN_Date_Conversion_edifact(ByRef Cell_Dt As String) As Object

        Try
            Dim T_Month, T_Date, T_Year As String

            Cell_Dt = Replace(Cell_Dt, "'", "")
            If Len(Cell_Dt) >= 5 Then
                T_Date = Mid(Cell_Dt, Len(Cell_Dt) - 1, 2)
                T_Month = Mid(Cell_Dt, Len(Cell_Dt) - 3, 2)
                T_Year = Mid(Cell_Dt, 1, Len(Cell_Dt) - 4)
                If Len(T_Year) = 1 Then
                    T_Year = "200" & T_Year
                ElseIf Len(T_Year) = 2 Then
                    T_Year = "20" & T_Year
                End If
                If IsDate(T_Date & "/" & T_Month & "/" & T_Year) = True Then
                    FN_Date_Conversion_edifact = T_Date & "/" & T_Month & "/" & T_Year
                Else
                    FN_Date_Conversion_edifact = ""
                End If
            Else
                FN_Date_Conversion_edifact = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Function FN_Display(ByVal TRANS_NUMBER As String, ByVal FileType As String) As String
        Dim Transit_Days, Row, Buffer_Days As Integer
        Dim adors As ADODB.Recordset 'sqlRdr

        Dim SCHQTY As Long
        Dim oCmd As ADODB.Command
        Dim SFTYDAYS_MNTD As Object
        Dim SAFETYDAYS_BELOW As Long
        Dim TMPWHSTOCK As Long
        Dim rsDate As ADODB.Recordset
        Dim sql As String = String.Empty, updSQL As String = String.Empty, strWHCode As String = String.Empty



        Dim varWhStock As Object = Nothing
        Dim varIssuedQty As Object = Nothing
        Dim varRcvdQty As Object = Nothing
        Dim varRevNo As Object = Nothing

        Dim dailypullrate As Long
        Dim intPos As Integer

        Dim varWHCODE As Object = Nothing
        Dim blnDAILYPULLFLAG As Integer
        Dim Rs As New ADODB.Recordset 'sqlRDr3


        Dim WHDATE As Date

        Dim sqlRdr As SqlDataReader ''adors
        Dim sqlRDr2 As SqlDataReader  '' rsTransitDays
        Dim sqlRDr3 As SqlDataReader  ''Rs
        Dim sqlRDr4 As SqlDataReader  ''RSWHSTOCK
        Dim sqlRDr5 As SqlDataReader  '' rsDate
        Dim sqlRDr6 As SqlDataReader  '' rsbagqty
        Dim sqlRDr7 As SqlDataReader  '' rsWHDt

        Try
            blnDAILYPULLFLAG = 0

            If FileType = "EDIFACT" Then
                sql = " Select Distinct D.Delivery_Dt,C.Cust_Drgno,I.Item_Code," & _
                    " I.DESCRIPTION,H.PARTY_ID1,H.PARTY_ID3,T.SAFETYSTOCK AS SAFETYSTKPERDAY," & _
                    " SP.SAFETYDAYSMAX,SP.SAFETYDAYS ,SP.SAFETYDAYSMIN,QUANTITY AS SHIPQTY," & _
                    " D.FREQUENCY,D.Dispatch_Pattern,D.Ran_No,T.StockCalcWAdays,T.ScheduleCalcMonths,T.DAYSFORSAFETYSTOCK,T.SUMOFRELEASEQTY  "
                sql = sql & " From SCHEDULE_UPLOAD_DARWINEDIFACT_DTL D with (nolock), CUSTITEM_MST C with (nolock),ITEM_MST I with (nolock), TMPSCHEDULESAFETYSTOCK T with (nolock), " & _
                    "SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H with (nolock), SCHEDULEPARAMETER_DTL SP with (nolock)"
                sql = sql & " Where C.Account_code=SP.Consignee_code and C.Cust_drgno= T.CUSTDRG_NO and C.Active=1"
                sql = sql & " And H.Doc_No=" & TRANS_NUMBER & ""
                sql = sql & " AND D.ITEM_CODE = C.CUST_DRGNO AND C.ITEM_CODE = I.ITEM_CODE"
                sql = sql & " AND D.UNIT_CODE = C.UNIT_CODE "
                sql = sql & " AND I.UNIT_CODE = C.UNIT_CODE "
                sql = sql & " AND T.UNIT_CODE = I.UNIT_CODE "
                sql = sql & " AND I.UNIT_CODE = H.UNIT_CODE "
                sql = sql & " AND H.UNIT_CODE = SP.UNIT_CODE and SP.UNIT_CODE='" & gstrUnitId & "'"
                sql = sql & " AND T.CUSTDRG_NO = C.CUST_DRGNO "
                sql = sql & " AND D.DOC_NO = H.DOC_NO " & _
                            " AND SP.CUSTOMER_CODE  = '" & Me.txtCustomerCode.Text & "' " & _
                            " AND SP.CUST_DRGNO = T.CUSTDRG_NO  " & _
                            " AND SP.WH_CODE = H.PARTY_ID1 " & _
                            " AND T.WH_CODE = H.PARTY_ID1 "
                sql = sql & " Order By D.DELIVERY_DT "
            End If

            If FileType = "COVISINT" Then
                sql = "Select Distinct D.Delivery_DATE,C.Cust_Drgno,I.Item_Code," & _
                    " I.DESCRIPTION,D.WH_CODE,T.SAFETYSTOCK AS SAFETYSTKPERDAY," & _
                    " SP.SAFETYDAYSMAX,SP.SAFETYDAYS  ,SP.SAFETYDAYSMIN,D.QTY AS SHIPQTY," & _
                    " D.FACTORY_CODE,D.CONSIGNEE_CODE,T.StockCalcWAdays,T.ScheduleCalcMonths,T.DAYSFORSAFETYSTOCK,T.SUMOFRELEASEQTY  "
                sql = sql & " From SCHEDULE_UPLOAD_COVISINT_DTL D with (nolock), CUSTITEM_MST C with (nolock),ITEM_MST I with (nolock), TMPSCHEDULESAFETYSTOCK T with (nolock), " & _
                    "SCHEDULE_UPLOAD_COVISINT_HDR H with (nolock), SCHEDULEPARAMETER_DTL SP with (nolock)"
                sql = sql & " Where C.Account_code=SP.Consignee_code and C.Cust_drgno= T.CUSTDRG_NO and C.Active=1"
                sql = sql & " And H.Doc_No=" & TRANS_NUMBER & ""
                sql = sql & " AND D.ITEM_CODE = C.CUST_DRGNO AND C.ITEM_CODE = I.ITEM_CODE"
                sql = sql & " AND T.CUSTDRG_NO = C.CUST_DRGNO "
                sql = sql & " AND D.DOC_NO = H.DOC_NO " & _
                            " AND SP.CUSTOMER_CODE  = '" & Me.txtCustomerCode.Text & "' " & _
                            " AND SP.CUST_DRGNO = T.CUSTDRG_NO " & _
                            " AND SP.WH_CODE = D.WH_CODE " & _
                            " AND T.WH_CODE = SP.WH_CODE "

                sql = sql & " AND D.UNIT_CODE = C.UNIT_CODE "
                sql = sql & " AND I.UNIT_CODE = C.UNIT_CODE "
                sql = sql & " AND T.UNIT_CODE = I.UNIT_CODE "
                sql = sql & " AND I.UNIT_CODE = H.UNIT_CODE "
                sql = sql & " AND H.UNIT_CODE = SP.UNIT_CODE and SP.UNIT_CODE='" & gstrUnitId & "'"
                sql = sql & " Order By D.DELIVERY_DATE "
            End If

            sqlRdr = SqlConnectionclass.ExecuteReader(sql)

            Row = 0 : spdRelease.MaxRows = Row
            SCHQTY = 0
            While sqlRdr.Read()

                mlngBAGQTY = 1

                sql = ""
                sql = " Select TransitDaysBySea, BufferDays "
                sql = sql & "  From ScheduleParameter_mst"
                sql = sql & " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and consignee_code = '" & sqlRdr("CONSIGNEE_CODE").ToString & "' and UNIT_CODE='" & gstrUnitId & "'"

                sqlRDr2 = SqlConnectionclass.ExecuteReader(sql)

                If sqlRDr2.HasRows Then
                    Transit_Days = IIf(IsDBNull(sqlRDr2.Item("TransitDaysBySea").ToString), 0, sqlRDr2.Item("TransitDaysBySea").ToString)
                    Buffer_Days = IIf(IsDBNull(sqlRDr2.Item("BufferDays").ToString), 0, sqlRDr2.Item("BufferDays").ToString)
                End If

                sqlRDr2.Close()

                If FileType = "EDIFACT" Then
                    If sqlRdr("Frequency").ToString = "" And sqlRdr("DISPATCH_PATTERN").ToString = "" Then
                        GoTo SKIP
                    End If
                End If
                oCmd = New ADODB.Command

                With oCmd
                    .ActiveConnection = mP_Connection
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc

                    If FileType = "EDIFACT" Then
                        .CommandText = "SCHEDULEQTY_EDIFACT"
                        .CommandTimeout = 0
                    End If

                    If FileType = "COVISINT" Then
                        .CommandText = "SCHEDULEQTY_COVISINT"
                        .CommandTimeout = 0
                    End If
                    .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, gstrUnitId))
                    .Parameters.Append(.CreateParameter("@CUST_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtCustomerCode.Text)))
                    .Parameters.Append(.CreateParameter("@DOCNO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , Trim(TRANS_NUMBER)))
                    .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(sqlRdr("Cust_DrgNo").ToString)))

                    If FileType = "EDIFACT" Then
                        .Parameters.Append(.CreateParameter("@WH_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 12, sqlRdr("PARTY_ID1").ToString))
                        .Parameters.Append(.CreateParameter("@TRANS_DT", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , Format(Trim(sqlRdr("DELIVERY_DT").ToString), "dd MMM yyyy")))
                        .Parameters.Append(.CreateParameter("@FACTORYCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, sqlRdr("PARTY_ID3").ToString))
                        'Rs.Open("SELECT CUSTOMER_CODE FROM CUSTOMER_MST" & _
                        '    " WHERE CUST_VENDOR_CODE = '" & sqlrdr("PARTY_ID1").tostring & "'" & _
                        '    " AND DOCK_CODE = '" & sqlrdr("PARTY_ID3").tostring & "' AND UNIT_CODE= '" & gstrUNITID & "'")
                        sql = ""
                        sql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST" & _
                            " WHERE CUST_VENDOR_CODE = '" & sqlRdr("PARTY_ID1").ToString & "'" & _
                            " AND DOCK_CODE = '" & sqlRdr("PARTY_ID3").ToString & "' AND UNIT_CODE= '" & gstrUnitId & "'"
                        sqlRDr3 = SqlConnectionclass.ExecuteReader(sql)
                        .Parameters.Append(.CreateParameter("@CONSIGNEE_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(sqlRDr3.Item("CUSTOMER_CODE").ToString)))
                    End If
                    If FileType = "COVISINT" Then
                        .Parameters.Append(.CreateParameter("@WH_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 12, sqlRdr("wh_code").ToString))
                        .Parameters.Append(.CreateParameter("@TRANS_DT", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , Trim(sqlRdr("DELIVERY_DATE").ToString)))
                        .Parameters.Append(.CreateParameter("@FACTORYCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(sqlRdr("Factory_Code").ToString)))
                        .Parameters.Append(.CreateParameter("@CONSIGNEE_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(sqlRdr("CONSIGNEE_CODE").ToString)))
                    End If

                    .Parameters.Append(.CreateParameter("@SCHQTY", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput))
                    .Execute()

                End With

                sql = ""
                sql = "SELECT WHSTOCK,MINSTOCK,RCVDQTY,ISSUEDQTY,WHDATE,REVNO FROM TMPWHSTOCK " & _
                         " WHERE ITEMCODE = '" & sqlRdr("Cust_DrgNo").ToString & "' " & _
                         " AND CUST_CODE = '" & txtCustomerCode.Text & "'  AND UNIT_CODE= '" & gstrUnitId & "'"

                If FileType = "EDIFACT" Then
                    sql = sql + " AND WH_CODE = '" & sqlRdr("PARTY_ID1").ToString & "' "
                    varWHCODE = sqlRdr("PARTY_ID1").ToString
                End If

                If FileType = "COVISINT" Then
                    sql = sql + " AND WH_CODE = '" & sqlRdr("wh_code").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdr("Factory_Code").ToString) & "'"
                    varWHCODE = sqlRdr("wh_code").ToString
                End If

                sqlRDr4 = SqlConnectionclass.ExecuteReader(sql)

                If sqlRDr4.HasRows Then
                    varWhStock = sqlRDr4.Item("WHSTOCK")
                    varIssuedQty = sqlRDr4.Item("ISSUEDQTY")
                    varRcvdQty = sqlRDr4.Item("RCVDQTY")
                    varRevNo = sqlRDr4.Item("REVNO")
                Else
                    varWhStock = 0
                    varIssuedQty = 0
                    varRcvdQty = 0
                    varRevNo = 0
                End If

                mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                sql = ""
                sql = "set dateformat 'dmy' select max(dt) as dt from Calendar_Mfg_mst" & _
                    " where work_flg=0 and unit_code='" & gstrUnitId & "' and "

                If FileType = "EDIFACT" Then
                    sql = sql + " dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), Format(sqlRdr("DELIVERY_DT").ToString, "dd MMM yyyy")) & "' "
                End If

                If FileType = "COVISINT" Then
                    sql = sql + " dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), Format(sqlRdr("DELIVERY_DATE").ToString, "dd MMM yyyy")) & "' "
                End If



                sqlRDr5 = SqlConnectionclass.ExecuteReader(sql)

                If DateDiff(Microsoft.VisualBasic.DateInterval.Day, sqlRDr5.Item("dt").value, ServerDate()) > 0 Then
                    GoTo NOSAFETYSTOCKPERDAY
                End If

                sql = ""
                sql = "select bag_qty from item_mst where item_code = '" & sqlRdr("Item_Code").ToString & "' and status = 'A'  AND UNIT_CODE= '" & gstrUnitId & "'"
                sqlRDr6 = SqlConnectionclass.ExecuteReader(sql)

                If sqlRDr6.HasRows Then
                    mlngBAGQTY = sqlRDr6.Item("bag_qty")
                Else
                    mlngBAGQTY = 1
                End If


                If sqlRDr4.HasRows Then
                    SCHQTY = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)
                Else
                    SCHQTY = IIf(IsDBNull(oCmd.Parameters("@SCHQTY").ToString), 0, oCmd.Parameters("@SCHQTY").ToString)
                End If

                If chkDlyPullQty.Checked = True Then

                    Dim schqty1 As String

                    sql = ""
                    sql = "select bag_qty from item_mst where item_code = '" & sqlRdr("Item_Code").ToString & "' and STATUS = 'A' AND UNIT_CODE= '" & gstrUnitId & "'"
                    sqlRDr6 = SqlConnectionclass.ExecuteReader(sql)

                    If sqlRDr6.HasRows Then
                        mlngBAGQTY = sqlRDr6.Item("bag_qty")
                    Else
                        mlngBAGQTY = 1
                    End If

                    schqty1 = FN_DAILYPULLQTY(varWhStock, varIssuedQty, varRcvdQty, SCHQTY, varWHCODE, sqlRdr("SAFETYDAYS").ToString, CStr(sqlRdr("Item_Code").ToString), CStr(sqlRdr("Cust_DrgNo").ToString), CInt(Row), CStr(sqlRdr("Factory_Code").ToString), mlngBAGQTY)
                    intPos = InStr(1, schqty1, "*")
                    dailypullrate = 0
                    If intPos > 0 Then
                        dailypullrate = Mid(schqty1, intPos + 1, Len(schqty1))
                        SCHQTY = Mid(schqty1, 1, intPos - 1)
                    End If
                    blnDAILYPULLFLAG = 1
                    If SCHQTY > 0 Then
                        Row = Row + 1 : Me.spdRelease.MaxRows = Row

                        If FileType = "EDIFACT" Then
                            Me.spdRelease.SetText(Enum_Up.Del_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                            Me.spdRelease.SetText(Enum_Up.Cust_Drg_No, Row, IIf(IsDBNull(sqlRdr("Cust_DrgNo")), "", sqlRdr("Cust_DrgNo")))
                            Me.spdRelease.SetText(Enum_Up.Item_Code, Row, IIf(IsDBNull(sqlRdr("Item_Code")), "", sqlRdr("Item_Code")))
                            Me.spdRelease.SetText(Enum_Up.Item_Desc, Row, IIf(IsDBNull(sqlRdr("Description")), "", sqlRdr("Description")))
                            Me.spdRelease.SetText(Enum_Up.SftyDays, Row, IIf(IsDBNull(sqlRdr("SafetyDays")), "", sqlRdr("SafetyDays")))
                            Me.spdRelease.SetText(Enum_Up.SftyStkPerDay, Row, IIf(IsDBNull(dailypullrate), "", dailypullrate))
                            Me.spdRelease.Row = Row
                            Me.spdRelease.Col = Enum_Up.Sch_Qty
                            Me.spdRelease.Text = IIf(IsDBNull(SCHQTY), 0, SCHQTY)
                            Me.spdRelease.SetText(Enum_Up.Hidd_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                            Me.spdRelease.SetText(Enum_Up.wh_code, Row, IIf(IsDBNull(sqlRdr("PARTY_ID1")), "", sqlRdr("PARTY_ID1")))

                            Me.spdRelease.SetText(Enum_Up.CONSIGNEE_CODE, Row, IIf(IsDBNull(sqlRdr("CONSIGNEE_CODE")), "", sqlRDr3.Item("CUSTOMER_CODE")))
                            Me.spdRelease.SetText(Enum_Up.RAN_No, Row, IIf(IsDBNull(sqlRdr("RAN_NO")), "", sqlRdr("RAN_NO")))
                        End If

                        If FileType = "COVISINT" Then
                            Me.spdRelease.SetText(Enum_Up.Del_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                            Me.spdRelease.SetText(Enum_Up.Cust_Drg_No, Row, IIf(IsDBNull(sqlRdr("Cust_DrgNo")), "", sqlRdr("Cust_DrgNo")))
                            Me.spdRelease.SetText(Enum_Up.Item_Code, Row, IIf(IsDBNull(sqlRdr("Item_Code")), "", sqlRdr("Item_Code")))
                            Me.spdRelease.SetText(Enum_Up.Item_Desc, Row, IIf(IsDBNull(sqlRdr("Description")), "", sqlRdr("Description")))
                            Me.spdRelease.SetText(Enum_Up.SftyDays, Row, IIf(IsDBNull(sqlRdr("SafetyDays")), "", sqlRdr("SafetyDays")))

                            Me.spdRelease.SetText(Enum_Up.SftyStkPerDay, Row, IIf(IsDBNull(dailypullrate), "", dailypullrate))
                            Me.spdRelease.Row = Row
                            Me.spdRelease.Col = Enum_Up.Sch_Qty
                            Me.spdRelease.Text = IIf(IsDBNull(SCHQTY), 0, SCHQTY)
                            Me.spdRelease.SetText(Enum_Up.Hidd_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                            Me.spdRelease.SetText(Enum_Up.wh_code, Row, IIf(IsDBNull(sqlRdr("wh_code")), "", sqlRdr("wh_code")))
                            Me.spdRelease.SetText(Enum_Up.CONSIGNEE_CODE, Row, IIf(IsDBNull(sqlRdr("CONSIGNEE_CODE")), "", sqlRdr("CONSIGNEE_CODE")))
                        End If
                    End If

                Else
                    sql = ""
                    sql = "select bag_qty from item_mst where item_code = '" & sqlRdr("Item_Code").ToString & "' and active_flg = 1 AND UNIT_CODE= '" & gstrUnitId & "'"
                    sqlRDr6 = SqlConnectionclass.ExecuteReader(sql)


                    If sqlRDr6.HasRows Then
                        mlngBAGQTY = sqlRDr6.Item("BAG_QTY")
                    Else
                        mlngBAGQTY = 1
                    End If

                    If sqlRdr("SAFETYSTKPERDAY").ToString > 0 Then

                        If ((sqlRdr("SAFETYSTKPERDAY").ToString * sqlRdr("safetydaysmin").ToString) - SCHQTY) <= 0 Then
                            SCHQTY = 0
                        Else
                            SAFETYDAYS_BELOW = Val(((sqlRdr("SAFETYSTKPERDAY").ToString * sqlRdr("safetydaysmin").ToString) - SCHQTY)) ''/ val(adors!SAFETYSTKPERDAY)
                            SFTYDAYS_MNTD = sqlRdr("safetydaysMAX").ToString - sqlRdr("safetydaysmin").ToString  ''+ SAFETYDAYS_BELOW
                            SFTYDAYS_MNTD = System.Math.Round(SFTYDAYS_MNTD, 0)

                            SCHQTY = (sqlRdr("SAFETYSTKPERDAY").ToString * SFTYDAYS_MNTD) + SAFETYDAYS_BELOW

                            SCHQTY = System.Math.Round(SCHQTY, 0)

                            Row = Row + 1 : Me.spdRelease.MaxRows = Row

                            If Val(mlngBAGQTY) >= Val(SCHQTY) Then
                                SCHQTY = mlngBAGQTY
                            Else
                                If mlngBAGQTY > 0 Then
                                    If SCHQTY Mod mlngBAGQTY > 0 Then
                                        SCHQTY = (mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)) + mlngBAGQTY
                                    Else
                                        SCHQTY = mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)
                                    End If
                                End If
                            End If

                            If FileType = "EDIFACT" Then
                                Me.spdRelease.SetText(Enum_Up.Del_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", VB6.Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.Cust_Drg_No, Row, IIf(IsDBNull(sqlRdr("Cust_DrgNo").ToString), "", sqlRdr("Cust_DrgNo").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Code, Row, IIf(IsDBNull(sqlRdr("Item_Code").ToString), "", sqlRdr("Item_Code").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Desc, Row, IIf(IsDBNull(sqlRdr("Description").ToString), "", sqlRdr("Description").ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyDays, Row, IIf(IsDBNull(SFTYDAYS_MNTD), "", SFTYDAYS_MNTD.ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyStkPerDay, Row, IIf(IsDBNull(sqlRdr("SAFETYSTKPERDAY").VALUE), "", sqlRdr("SAFETYSTKPERDAY").VALUE.ToString))

                                Me.spdRelease.Row = Row
                                Me.spdRelease.Col = Enum_Up.Sch_Qty
                                Me.spdRelease.Text = IIf(IsDBNull(SCHQTY), 0, SCHQTY)

                                Me.spdRelease.SetText(Enum_Up.Hidd_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", VB6.Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.wh_code, Row, IIf(IsDBNull(sqlRdr("PARTY_ID1").ToString), "", sqlRdr("PARTY_ID1").ToString))
                                Me.spdRelease.SetText(Enum_Up.RAN_No, Row, IIf(IsDBNull(sqlRdr("RAN_NO")), "", sqlRdr("RAN_NO")))
                            End If

                            If FileType = "COVISINT" Then
                                Me.spdRelease.SetText(Enum_Up.Del_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", VB6.Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.Cust_Drg_No, Row, IIf(IsDBNull(sqlRdr("Cust_DrgNo").ToString), "", sqlRdr("Cust_DrgNo").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Code, Row, IIf(IsDBNull(sqlRdr("Item_Code").ToString), "", sqlRdr("Item_Code").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Desc, Row, IIf(IsDBNull(sqlRdr("Description").ToString), "", sqlRdr("Description").ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyDays, Row, IIf(IsDBNull(SFTYDAYS_MNTD), "", SFTYDAYS_MNTD.ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyStkPerDay, Row, IIf(IsDBNull(sqlRdr("SAFETYSTKPERDAY").VALUE), "", sqlRdr("SAFETYSTKPERDAY").VALUE.ToString))

                                Me.spdRelease.Row = Row
                                Me.spdRelease.Col = Enum_Up.Sch_Qty
                                Me.spdRelease.Text = IIf(IsDBNull(SCHQTY), 0, SCHQTY)

                                Me.spdRelease.SetText(Enum_Up.Hidd_Dt, Row, IIf(IsDBNull(sqlRDr5.Item("dt").ToString), "", VB6.Format(sqlRDr5.Item("dt").ToString, "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.wh_code, Row, IIf(IsDBNull(sqlRdr("wh_code").ToString), "", sqlRdr("wh_code").ToString))
                                Me.spdRelease.SetText(Enum_Up.CONSIGNEE_CODE, Row, IIf(IsDBNull(sqlRdr("CONSIGNEE_CODE").ToString), "", sqlRdr("CONSIGNEE_CODE").ToString))
                            End If
                        End If
                    Else
                        If SCHQTY > 0 Then

                            If Val(varWhStock) <> SCHQTY Then
                                SCHQTY = 0
                                GoTo NOSAFETYSTOCKPERDAY
                            Else
                                GoTo SKIP
                            End If
                        Else
                            SCHQTY = Val(varIssuedQty) - (Val(varRcvdQty) + Val(varWhStock))

                        End If

                        SFTYDAYS_MNTD = sqlRdr("SafetyDaysMax").ToString - sqlRdr("safetydaysmin").ToString ''+ SAFETYDAYS_BELOW
                        SFTYDAYS_MNTD = System.Math.Round(SFTYDAYS_MNTD, 0)
                        'mayur
                        If SCHQTY > 0 Then
                            If FileType = "EDIFACT" Then
                                Row = Row + 1
                                Me.spdRelease.MaxRows = Row
                                Me.spdRelease.SetText(Enum_Up.Del_Dt, Row, IIf(IsDBNull(sqlRdr("DELIVERY_DT").ToString), "", VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -(Transit_Days + Buffer_Days), sqlRdr("DELIVERY_DT").value), "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.Cust_Drg_No, Row, IIf(IsDBNull(sqlRdr("Cust_DrgNo").ToString), "", sqlRdr("Cust_DrgNo").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Code, Row, IIf(IsDBNull(sqlRdr("Item_Code").ToString), "", sqlRdr("Item_Code").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Desc, Row, IIf(IsDBNull(sqlRdr("Description").ToString), "", sqlRdr("Description").ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyDays, Row, IIf(IsDBNull(SFTYDAYS_MNTD), "", SFTYDAYS_MNTD.ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyStkPerDay, Row, IIf(IsDBNull(sqlRdr("SAFETYSTKPERDAY").VALUE), "", sqlRdr("SAFETYSTKPERDAY").VALUE.ToString))

                                Me.spdRelease.Row = Row
                                Me.spdRelease.Col = Enum_Up.Sch_Qty
                                Me.spdRelease.Text = IIf(IsDBNull(SCHQTY), 0, SCHQTY)

                                Me.spdRelease.SetText(Enum_Up.Hidd_Dt, Row, IIf(IsDBNull(sqlRdr("DELIVERY_DT").ToString), "", VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -Transit_Days, sqlRdr("DELIVERY_DT").value), "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.wh_code, Row, IIf(IsDBNull(sqlRdr("PARTY_ID1").ToString), "", sqlRdr("PARTY_ID1").ToString))
                                Me.spdRelease.SetText(Enum_Up.RAN_No, Row, IIf(IsDBNull(sqlRdr("RAN_NO")), "", sqlRdr("RAN_NO")))
                            End If

                            If FileType = "COVISINT" Then
                                Row = Row + 1
                                Me.spdRelease.MaxRows = Row
                                Me.spdRelease.SetText(Enum_Up.Del_Dt, Row, IIf(IsDBNull(sqlRdr("DELIVERY_DATE").ToString), "", VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -(Transit_Days + Buffer_Days), sqlRdr("DELIVERY_DATE").value), "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.Cust_Drg_No, Row, IIf(IsDBNull(sqlRdr("Cust_DrgNo").ToString), "", sqlRdr("Cust_DrgNo").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Code, Row, IIf(IsDBNull(sqlRdr("Item_Code").ToString), "", sqlRdr("Item_Code").ToString))
                                Me.spdRelease.SetText(Enum_Up.Item_Desc, Row, IIf(IsDBNull(sqlRdr("Description").ToString), "", sqlRdr("Description").ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyDays, Row, IIf(IsDBNull(SFTYDAYS_MNTD), "", SFTYDAYS_MNTD.ToString))
                                Me.spdRelease.SetText(Enum_Up.SftyStkPerDay, Row, IIf(IsDBNull(sqlRdr("SAFETYSTKPERDAY").VALUE), "", sqlRdr("SAFETYSTKPERDAY").VALUE.ToString))

                                Me.spdRelease.Row = Row
                                Me.spdRelease.Col = Enum_Up.Sch_Qty
                                Me.spdRelease.Text = IIf(IsDBNull(SCHQTY), 0, SCHQTY)

                                Me.spdRelease.SetText(Enum_Up.Hidd_Dt, Row, IIf(IsDBNull(sqlRdr("DELIVERY_DATE").ToString), "", VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -Transit_Days, sqlRdr("DELIVERY_DATE").value), "dd/MM/yyyy")))
                                Me.spdRelease.SetText(Enum_Up.wh_code, Row, IIf(IsDBNull(sqlRdr("wh_code").ToString), "", sqlRdr("wh_code").ToString))
                                Me.spdRelease.SetText(Enum_Up.CONSIGNEE_CODE, Row, IIf(IsDBNull(sqlRdr("CONSIGNEE_CODE").ToString), "", sqlRdr("CONSIGNEE_CODE").ToString))
                            End If
                        End If

                    End If
                End If

                oCmd = Nothing

NOSAFETYSTOCKPERDAY:

                sql = ""
                sql = "select top 1 trans_dt,customer_code," & _
                    " revno from warehouse_stock_dtl" & _
                    " where customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE= '" & gstrUnitId & "'"

                If FileType = "EDIFACT" Then
                    sql = sql & " and  warehouse_code = '" & sqlRdr("PARTY_ID1").ToString & "'"
                End If

                If FileType = "COVISINT" Then
                    sql = sql + " and  warehouse_code = '" & sqlRdr("wh_code").ToString & "'"
                End If

                sql = sql + " group by customer_code, trans_dt,revno" & _
                " order by trans_dt desc,revno desc "


                sqlRDr7 = SqlConnectionclass.ExecuteReader(sql)

                If sqlRDr7.HasRows Then
                    WHDATE = sqlRDr7.Item("TRANS_DT")
                Else
                    WHDATE = ""
                End If
                If chkDlyPullQty.Checked = True Then
                    blnDAILYPULLFLAG = 1
                Else
                    blnDAILYPULLFLAG = 0
                End If

                'mayur

                If FileType = "EDIFACT" Then
                    sql = ""
                    sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                        " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                        " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                        " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
                        " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                        " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
                        " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdr("PARTY_ID1").ToString & "', " & _
                        " '" & Trim(sqlRdr("Cust_DrgNo").ToString) & "','" & Format(sqlRdr("DELIVERY_DT").ToString, "dd MMM yyyy") & "'," & _
                        " '" & Val(sqlRdr("shipqty").ToString) & "','" & Format(sqlRDr5.Item("dt").ToString, "dd MMM yyyy") & "'," & _
                        " '" & SCHQTY & "','" & sqlRDr3.Item("CUSTOMER_CODE").ToString & "','" & IIf(IsDBNull(varWhStock), 0, varWhStock) & "'," & _
                        " '" & varRcvdQty & "','" & varIssuedQty & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                        " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                        " " & sqlRdr("safetydaysmin").ToString & "," & sqlRdr("safetydaysMAX").ToString & "," & _
                        " " & sqlRdr("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                        " '" & sqlRdr("STOCKCALCWADAYS").ToString & "','" & sqlRdr("ScheduleCalcMonths").ToString & "'" & _
                        " ,'" & sqlRdr("DAYSFORSAFETYSTOCK").ToString & "'," & dailypullrate & ",'" & sqlRdr("SUMOFRELEASEQTY").ToString & "'," & Transit_Days & "," & _
                        " " & Buffer_Days & " ,'" & gstrUnitId & "')"
                    SqlConnectionclass.ExecuteNonQuery(sql)
                End If

                If FileType = "COVISINT" Then
                    sql = ""
                    sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                        " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                        " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                        " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG" & _
                        " ,SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                        " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY, transitDays,bufferDays,UNIT_CODE)" & _
                        " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdr("wh_code").ToString & "', " & _
                        " '" & Trim(sqlRdr("Cust_DrgNo").ToString) & "','" & Format(sqlRdr("DELIVERY_DATE").ToString, "dd MMM yyyy") & "'," & _
                        " '" & Val(sqlRdr("shipqty").ToString) & "','" & Format(sqlRDr5.Item("dt").ToString, "dd MMM yyy") & "'," & _
                        " '" & SCHQTY & "','" & sqlRdr("CONSIGNEE_CODE").ToString & "','" & varWhStock & "'," & _
                        " '" & varRcvdQty & "','" & varIssuedQty & "',getDate(),'" & mP_User & "' ,getDate(),'" & mP_User & "'," & _
                        " '" & Format(WHDATE, "dd MMM yyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                        " " & sqlRdr("safetydaysmin").ToString & "," & sqlRdr("safetydaysMAX").ToString & "," & _
                        " " & sqlRdr("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                        " '" & sqlRdr("STOCKCALCWADAYS").ToString & "','" & sqlRdr("ScheduleCalcMonths").ToString & "'" & _
                        " ,'" & sqlRdr("DAYSFORSAFETYSTOCK").ToString & "'," & dailypullrate & ",'" & sqlRdr("SUMOFRELEASEQTY").ToString & "'," & Transit_Days & "," & _
                        " " & Buffer_Days & ",'" & gstrUnitId & "' )"
                    SqlConnectionclass.ExecuteNonQuery(sql)
                End If

                '''''''''end

                TMPWHSTOCK = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)

                If Val(sqlRdr("shipqty").ToString) >= 0 Then
                    TMPWHSTOCK = TMPWHSTOCK + SCHQTY                        ''- val(adors!SHIPQTY)
                End If
                updSQL = "UPDATE TMPWHSTOCK" & _
                    " SET WHSTOCK = " & TMPWHSTOCK & " , "

                If FileType = "EDIFACT" Then
                    updSQL = updSQL + " WHDATE = '" & Format(sqlRdr("DELIVERY_DT").ToString, "dd MMM yyyy") & "'"
                End If

                If FileType = "COVISINT" Then
                    updSQL = updSQL + " WHDATE = '" & Format(sqlRdr("DELIVERY_DATE").ToString, "dd MMM yyyy") & "'"
                End If

                updSQL = updSQL + " WHERE ITEMCODE = '" & sqlRdr("Cust_DrgNo").ToString & "' " & _
                    " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE= '" & gstrUnitId & "'"

                If FileType = "EDIFACT" Then
                    updSQL = updSQL + " AND WH_CODE = '" & sqlRdr("PARTY_ID1").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdr("PARTY_ID3").ToString) & "'"
                End If

                If FileType = "COVISINT" Then
                    updSQL = updSQL + " AND WH_CODE = '" & sqlRdr("wh_code").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdr("Factory_Code").ToString) & "'"
                End If

                SqlConnectionclass.ExecuteNonQuery(updSQL)
SKIP:
                sqlRdr.NextResult()
            End While

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            sqlRdr.Close() : sqlRdr = Nothing
            sqlRDr2.Close() : sqlRDr2 = Nothing
            sqlRDr3.Close() : sqlRDr3 = Nothing
            sqlRDr4.Close() : sqlRDr4 = Nothing
            sqlRDr5.Close() : sqlRDr5 = Nothing
            sqlRDr6.Close() : sqlRDr6 = Nothing
            sqlRDr7.Close() : sqlRDr7 = Nothing
        End Try
    End Function

    Private Function PopulateWHDtls(ByRef RevNo As Short) As Object
        Try
            Dim sql As String = String.Empty
            Dim Row As Short
            Dim cmdproc As New ADODB.Command
            Dim sqlRdr As SqlDataReader


            sql = "Sp_CalculateSafetyStockforWH  '" & gstrUnitId & "', '" & Trim(txtCustomerCode.Text) & "','" & Trim(txtUnitCode.Text) & "','" & Trim(txtConsignee.Text) & "','W',0,'" & DTPicker1.Value & "','" & gstrIpaddressWinSck & "'," & RevNo
            cmdproc.let_ActiveConnection(mP_Connection)
            cmdproc.CommandTimeout = 0
            cmdproc.CommandText = sql
            cmdproc.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            cmdproc = Nothing

            sql = ""
            sql = "select Cust_code,Consignee_code,WH_Code,Custdrg_no,tmpWarehouseSafetyStock.Item_code," & " WarehouseStock , SafetyStock, StockDate, Description " & " From tmpWarehouseSafetyStock, item_mst " & " where tmpWarehouseSafetyStock.UNIT_CODE=item_mst.UNIT_CODE and tmpWarehouseSafetyStock.item_code=item_mst.item_code and ip_address='" & gstrIpaddressWinSck & "' and UNIT_CODE='" & gstrUnitId & "'"
            sqlRdr = SqlConnectionclass.ExecuteReader(sql)

            Row = 1 : Me.spdWareHouse.MaxRows = 1
            While sqlRdr.Read()

                Me.spdWareHouse.SetText(enumWH.Stock_dt, Row, IIf(IsDBNull(sqlRdr("StockDate").ToString), "", VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 0, sqlRdr("StockDate").value), "dd/MM/yyyy")))
                Me.spdWareHouse.SetText(enumWH.CustPartNo, Row, IIf(IsDBNull(sqlRdr("CUSTDRG_NO").ToString), "", sqlRdr("CUSTDRG_NO").ToString))
                Me.spdWareHouse.SetText(enumWH.ItemCode, Row, IIf(IsDBNull(sqlRdr("Item_Code").ToString), "", sqlRdr("Item_Code").ToString))
                Me.spdWareHouse.SetText(enumWH.Description, Row, IIf(IsDBNull(sqlRdr("Description").ToString), "", sqlRdr("Description").ToString))
                Me.spdWareHouse.SetText(enumWH.AvlStk, Row, IIf(IsDBNull(sqlRdr("warehousestock").ToString), "0", sqlRdr("warehousestock").ToString))
                Me.spdWareHouse.SetText(enumWH.SaftyStk, Row, IIf(IsDBNull(sqlRdr("SafetyStock").ToString), "0", sqlRdr("SafetyStock").ToString))

                If sqlRdr("warehousestock").ToString < sqlRdr("SafetyStock").ToString Then
                    With spdWareHouse
                        .Col = enumWH.Stock_dt
                        .Col2 = enumWH.SaftyStk
                        .Row = Row
                        .BlockMode = True
                        .BackColor = System.Drawing.Color.Red
                        .BlockMode = False
                    End With
                End If

                Me.spdWareHouse.MaxRows = Me.spdWareHouse.MaxRows + 1
                Row = Row + 1
                sqlRdr.NextResult()
            End While

            Me.spdWareHouse.MaxRows = Me.spdWareHouse.MaxRows - 1

            sqlRdr.Close() : sqlRdr = Nothing
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
        Return Nothing
    End Function

    Private Function FN_Find_Revision() As Short
        Try
            Dim sql As String
            Dim sqlRdr As SqlDataReader

            If Me.OptReleaseFile.Checked = True Then
                sql = " Select max(RevNo) + 1 as RevNO"
                sql = sql & " From Schedule_Upload_Darwin_Hdr"
                sql = sql & " Where Cust_Code ='" & Trim(txtCustomerCode.Text) & "' and consignee_code = '" & txtConsignee.Text & "'"
                sql = sql & " And Ent_dt=getdate()"
                sql = sql & " And UNIT_CODE='" & gstrUnitId & "'"
            Else
                sql = " Select max(RevNo) + 1 as RevNO"
                sql = sql & " From WareHouse_Stock_Dtl with (nolock)"
                sql = sql & " Where Customer_Code ='" & Trim(txtCustomerCode.Text) & "'"
                sql = sql & " and Warehouse_Code = '" & Trim(Me.txtUnitCode.Text) & "' and consignee_code='" & Trim(txtConsignee.Text) & "'"
                sql = sql & " and trans_dt = '" & VB6.Format(Trim(Me.DTPicker1.Value), "DD MMM YYYY") & "'"
                sql = sql & " And UNIT_CODE='" & gstrUnitId & "'"
            End If

            sqlRdr = SqlConnectionclass.ExecuteReader(sql)
            If sqlRdr.HasRows Then
                While sqlRdr.Read
                    FN_Find_Revision = IIf(IsDBNull(sqlRdr("RevNo").ToString), 1, sqlRdr("RevNo").ToString)
                End While
            Else
                FN_Find_Revision = 1
            End If
            sqlRdr.Close()
            sqlRdr = Nothing

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Function

    Private Sub FN_Spread_Settings()

        Try
            chkDlyPullQty.Visible = False
            With Me.spdRelease
                .MaxCols = 11 : .MaxRows = 0

                .Row = 0
                .Col = Enum_Up.Del_Dt : .Text = "Shipment Date "
                .Col = Enum_Up.Cust_Drg_No : .Text = "Customer Drawing No"
                .Col = Enum_Up.Item_Code : .Text = "Item Code "
                .Col = Enum_Up.Item_Desc : .Text = "Description "
                .Col = Enum_Up.SftyDays : .Text = "Safety Days"
                .Col = Enum_Up.SftyStkPerDay : .Text = "Safety Stock/Day"
                .Col = Enum_Up.Sch_Qty : .Text = " Shipping Qty "
                .Col = Enum_Up.Hidd_Dt : .Text = ""
                .Col = Enum_Up.wh_code : .Text = "WareHouse Code"
                .Col = Enum_Up.CONSIGNEE_CODE : .Text = "Consignee Code"
                .Col = Enum_Up.RAN_No : .Text = "RAN No."
                .set_ColWidth(Enum_Up.wh_code, 100)
                .Row = -1
                .Col = Enum_Up.Del_Dt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText 'eMpro-20080911-21453
                .Col = Enum_Up.Cust_Drg_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = Enum_Up.Item_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = Enum_Up.Item_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = Enum_Up.SftyDays : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = Enum_Up.SftyStkPerDay : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = Enum_Up.Sch_Qty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText 'eMpro-20080911-21453 
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = Enum_Up.Hidd_Dt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = Enum_Up.wh_code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = Enum_Up.CONSIGNEE_CODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = Enum_Up.RAN_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

                .set_ColWidth(1, 10)
                .set_ColWidth(2, 19)
                .set_ColWidth(3, 15)
                .set_ColWidth(4, 16)
                .set_ColWidth(5, 20)
                .set_ColWidth(6, 20)
                .set_ColWidth(7, 12)
                .set_ColWidth(Enum_Up.wh_code, 20)
                .set_ColWidth(Enum_Up.CONSIGNEE_CODE, 20)
                .set_ColWidth(8, 0)
                .set_ColWidth(Enum_Up.RAN_No, 0)
                .set_RowHeight(0, 15)

            End With

            With Me.spdWareHouse

                .MaxCols = 6 : .MaxRows = 0

                .Row = 0
                .Col = enumWH.Stock_dt : .Text = "Stock Date "
                .Col = enumWH.CustPartNo : .Text = " Customer Part No "
                .Col = enumWH.ItemCode : .Text = "Item Code "
                .Col = enumWH.Description : .Text = "Description "
                .Col = enumWH.AvlStk : .Text = "Stock Available"
                .Col = enumWH.SaftyStk : .Text = "Safety Stock "

                .Row = -1
                .Col = enumWH.Stock_dt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText 'eMpro-20080911-21453
                .Col = enumWH.CustPartNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enumWH.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enumWH.Description : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enumWH.AvlStk : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enumWH.SaftyStk : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .ColHidden = True

                .set_ColWidth(1, 10)
                .set_ColWidth(2, 19)
                .set_ColWidth(3, 15)
                .set_ColWidth(4, 16)
                .set_ColWidth(5, 12)
                .set_ColWidth(6, 12)
                .set_RowHeight(0, 15)

            End With

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Sub


    Private Sub FN_WareHouse_File_Upload()
        Dim Trans_Satus As Boolean
        Try
            Dim Item_Prefix, Item_Suffix As String
            Dim cmd As ADODB.Command
            Dim sql As String
            Dim Col, Row As Short

            Dim Rev_No As Short
            Dim lngStockQty As Integer

            Dim Msg As String
            Dim Flag As Short




            Dim sqlRdr As SqlDataReader 'rsItems
            Dim sqlRdr2 As SqlDataReader 'RSwh
            Dim strSql As String = String.Empty

            Dim Item_Rate As Double

            Dim WhStkObj As New prj_uploadInvoiceDaimler.prj_uploadInvoiceDaimler          'eMpro-20090309-28458

            If Me.txtFileName.Text.ToUpper = "Unknown".ToUpper Then
                MsgBox("Please Select the Upload File", MsgBoxStyle.Information, ResolveResString(100))
                txtFileName.Focus()
                Bool_Not_File = True
                Exit Sub
            End If

            Obj_FSO = New Scripting.FileSystemObject

            If Obj_FSO.FileExists(Me.txtFileName.Text) = False Then
                MsgBox(" File Does not Exist ", MsgBoxStyle.Information, ResolveResString(100))
                txtFileName.Focus()
                Bool_Not_File = True
                Exit Sub
            End If

            If UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) <> UCase("txt") Then
                Obj_EX = New Excel.Application
                Obj_EX.Workbooks.Open(Trim(Me.txtFileName.Text))
            End If

            If OptStock.Checked = True And UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) = UCase("txt") Then
                Rev_No = FN_Find_Revision()
                Msg = WhStkObj.FN_WareHouse_TextFileUpload(txtCustomerCode.Text, txtUnitCode.Text, txtConsignee.Text, DTPicker1.Value, gstrConnectSQLClient, txtFileName.Text, Rev_No, mP_User)
                MsgBox(Mid(Msg, 3, Msg.Length))
                If Mid(Msg, 1, 1) = "Y" Then
                    Call PopulateWHDtls(Rev_No)
                Else
                    spdWareHouse.MaxRows = 0
                End If
                Exit Sub
            End If

            If OptStock.Checked = True And UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) <> UCase("txt") Then
                If Not UCase(Obj_EX.Sheets(1).Name) = UCase("StockStatus") Then
                    MsgBox("Name Of Default Sheet Must be Stock Status")
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                Else
                    Obj_EX.Sheets.Item(1).Select()
                End If
            End If

            If OptRecvd.Checked = True And UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) <> UCase("txt") Then
                If Not UCase(Obj_EX.Sheets(2).Name) = UCase("Receiving") Then
                    MsgBox("Name Of Default Sheet Must be Receiving")

                    Obj_FSO = Nothing

                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                Else

                    Obj_EX.Sheets.Item(2).Select()
                End If
            End If
            Dim countREC As Short
            If OptStock.Checked = True Then
                Row = 4 : Col = 2

                If Obj_EX.Range("$B$" & Row).Value Is Nothing Then
                    Item_Suffix = ""
                Else
                    Item_Suffix = Obj_EX.Range("$B$" & Row).Value2.ToString
                End If

                Row = 3 : Col = 2

                If Obj_EX.Range("$B$" & Row).Value Is Nothing Then
                    Item_Prefix = ""
                Else
                    Item_Prefix = Obj_EX.Range("$B$" & Row).Value2.ToString
                End If

                Row = 5 : Col = 2

                If Obj_EX.Range("$B$" & Row).Value Is Nothing Then
                    lngStockQty = 0
                Else
                    lngStockQty = Convert.ToInt64(Obj_EX.Range("$B$" & Row).Value2.ToString)
                End If

                Row = 8 : Col = 2

                If Obj_EX.Range("$B$" & Row).Value Is Nothing Then
                    Item_Rate = 0.0
                Else
                    Item_Rate = Convert.ToDouble(Obj_EX.Range("$B$" & Row).Value2.ToString)
                End If

                If Len(Item_Suffix) = 0 Then

                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If

                    MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If

                Rev_No = FN_Find_Revision()

                mP_Connection.BeginTrans()
                Trans_Satus = True


                While Len(Item_Suffix) <> 0
                    Row = 3
                    strSql = ""
                    strSql = "INSERT INTO WAREHOUSE_STOCK_DTL(CUSTOMER_CODE, " & "WAREHOUSE_CODE,Consignee_Code,UPLOAD_FILE_NAME,ITEM_CODE,QTY,RATE, " & "TRANS_DT,ENT_DT,ENT_ID, REVNO ,UNIT_CODE)" & "VALUES ('" & Me.txtCustomerCode.Text & "','" & Me.txtUnitCode.Text & "'  ,'" & Me.txtConsignee.Text & "'," & " '" & Me.txtFileName.Text & "','" & Item_Prefix & "" & Item_Suffix & "' ," & " " & lngStockQty & "," & Item_Rate & ",'" & VB6.Format(DTPicker1.Value, "dd/MM/yyyy") & "', " & "  getDate() , '" & mP_User & "'," & Rev_No & " ,'" & gstrUnitId & "')"
                    SqlConnectionclass.ExecuteNonQuery(strSql)

                    Col = Col + 1
                    range = Obj_EX.Cells(Row, Col)
                    If Not range.Value Is Nothing Then
                        If Trim(range.Value.ToString) <> "" Then
                            Item_Prefix = range.Value.ToString
                        End If
                    End If

                    Row = 4

                    range = Obj_EX.Cells(Row, Col)

                    If Not range.Value Is Nothing Then
                        Item_Suffix = range.Value.ToString
                    Else
                        Item_Suffix = ""
                    End If
                    Row = 5

                    range = Obj_EX.Cells(Row, Col)
                    If Not range.Value Is Nothing Then
                        lngStockQty = Convert.ToInt64(range.Value.ToString)
                    Else
                        lngStockQty = 0
                    End If
                    Row = 8

                    range = Obj_EX.Cells(Row, Col)
                    If Not range.Value Is Nothing Then
                        Item_Rate = Convert.ToDouble(range.Value.ToString)
                    Else
                        Item_Rate = 0.0
                    End If

                End While

                strSql = ""
                strSql = "DELETE FROM WAREHOUSE_STOCK_DTL " & "WHERE ITEM_CODE NOT IN (SELECT CUST_DRGNO FROM CUSTITEM_MST " & "WHERE ACCOUNT_CODE = '" & txtConsignee.Text & "' " & "AND SCHUPLDREQD = 1 AND ACTIVE=1 and UNIT_CODE= '" & gstrUnitId & "' ) AND CUSTomer_CODE = '" & txtCustomerCode.Text & "' " & "and consignee_code = '" & txtConsignee.Text & "' and revno = " & Rev_No & " and UNIT_CODE= '" & gstrUnitId & "'"
                SqlConnectionclass.ExecuteNonQuery(strSql)

                strSql = ""
                strSql = "SELECT * FROM WAREHOUSE_STOCK_DTL " & " WHERE WAREHOUSE_CODE = '" & txtUnitCode.Text & "'" & " AND CUSTomer_CODE = '" & txtCustomerCode.Text & "' " & " AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' " & " AND revno = " & Rev_No & " and trans_dt = '" & DTPicker1.Value & "' and UNIT_CODE= '" & gstrUnitId & "'"

                sqlRdr2 = SqlConnectionclass.ExecuteReader(strSql)
                Trans_Satus = False
                'mayur

                sql = "select  ltrim(rtrim(item_code)) + '     ' + ltrim(rtrim(cust_drgno))" & _
                 " as item_custdrgno" & " From custitem_mst " & _
                 " where account_code = '" & txtConsignee.Text & "' " & " " & _
                 " and UNIT_CODE= '" & gstrUnitId & "' and cust_drgno in (select distinct item_code from warehouse_stock_dtl" & _
                 " with (nolock)" & " where revno = " & Rev_No & " and   UNIT_CODE= '" & gstrUnitId & "'" & _
                 " customer_code = '" & txtCustomerCode.Text & "' and" & "" & _
                 " consignee_code = '" & txtConsignee.Text & "' " & " and" & _
                 " trans_dt = '" & Me.DTPicker1.Value & "')" & " and" & _
                 " ltrim(rtrim(item_code)) + '     ' + ltrim(rtrim(cust_drgno))" & _
                 " NOT in " & " (select ltrim(rtrim(item_code)) + '     ' + ltrim(rtrim(cust_drgno))" & _
                 " as item_custdrgno " & " From custitem_mst where" & _
                 " account_code = '" & txtCustomerCode.Text & "' AND ACTIVE = 1 AND SCHUPLDREQD = 1  and UNIT_CODE= '" & gstrUnitId & "') AND ACTIVE = 1 AND SCHUPLDREQD = 1"


                sqlRdr = SqlConnectionclass.ExecuteReader(sql)
                Msg = ""
                If sqlRdr.HasRows Then
                    While sqlRdr.Read
                        Msg = Msg & "  " & vbCrLf + sqlRdr("item_custdrgno")
                        sqlRdr.NextResult()
                    End While


                    MsgBox("Following Items Are Not Defined In The Customer Item Master : " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    Flag = 1
                End If

                sql = " Select distinct W.item_code  from WAREHOUSE_STOCK_dtl W"
                sql = sql & " Where LTrim(RTrim(w.Item_Code))"
                sql = sql & " not in (select Cust_DrgNo  from ScheduleParameter_dtl where  Customer_code = '" & Trim(txtCustomerCode.Text) & "'  AND CONSIGNEE_CODE = '" & Trim(txtConsignee.Text) & "'  And WH_Code = '" & Me.txtUnitCode.Text & "'  and UNIT_CODE= '" & gstrUnitId & "')"
                sql = sql & " and W.revno = " & Rev_No & " and W.customer_code = '" & Trim(txtCustomerCode.Text) & "' AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' and w.warehouse_code = '" & Me.txtUnitCode.Text & "' and w.trans_dt = '" & Me.DTPicker1.Value & "' and w.UNIT_CODE= '" & gstrUnitId & "'"
                Msg = ""



                sqlRdr.Close()
                sqlRdr = Nothing

                sqlRdr = SqlConnectionclass.ExecuteReader(sql)

                If sqlRdr.FieldCount > 0 Then

                    While sqlRdr.Read
                        Msg = Msg & "  " & vbCrLf + sqlRdr("item_code")
                        sqlRdr.NextResult()
                    End While

                    MsgBox("These Items Are Not Defined In The Schedule Parameter.: " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    Flag = 1
                End If

                sql = "select c.account_code, c.cust_drgno, count(*) countitem " & " from custitem_mst c where Exists(   Select * " & " from warehouse_stock_dtl w Where w.Item_Code = c.Cust_DrgNo" & " and w.trans_dt = '" & Me.DTPicker1.Value & "' and w.revno = '" & Rev_No & "' " & " and c.account_code = w.CONSIGNEE_code  and UNIT_CODE= '" & gstrUnitId & "' ) and c.UNIT_CODE= '" & gstrUnitId & "' and     c.active = 1 AND SCHUPLDREQD = 1 and c.account_code = '" & Me.txtConsignee.Text & "'" & " group by account_code, cust_drgno "

                sqlRdr.Close()
                sqlRdr = Nothing

                sqlRdr = SqlConnectionclass.ExecuteReader(sql)

                Msg = ""

                If sqlRdr.FieldCount > 0 Then
                    While sqlRdr.Read
                        countREC = sqlRdr("COUNTITEM")
                        If countREC > 1 Then
                            Msg = Msg & vbCrLf + sqlRdr("cust_drgno")
                            Flag = 1
                        End If
                        sqlRdr.NextResult()
                    End While
                    If Msg <> "" Then
                        MsgBox("For Consignee: " & txtConsignee.Text & " Following Cust_DrgNo Are Active For Multiple Items : " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    End If
                End If

                If Flag = 1 Then
                    mP_Connection.RollbackTrans()
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If

                    Exit Sub
                Else
                    If sqlRdr.FieldCount <= 0 Then
                        mP_Connection.CommitTrans()
                        MsgBox("Warehouse Stock Not Uploaded As No Item Defined In The System", MsgBoxStyle.Information, ResolveResString(100))
                        sqlRdr.Close()
                        sqlRdr = Nothing
                    Else
                        mP_Connection.CommitTrans()
                        MsgBox("WareHouse Stock Uploaded Succesfully !", MsgBoxStyle.Information, ResolveResString(100))
                        Call PopulateWHDtls(Rev_No)

                        Obj_FSO = Nothing

                        If Not Obj_EX Is Nothing Then
                            KillExcelProcess(Obj_EX)
                            Obj_EX = Nothing
                        End If


                        sqlRdr.Close()
                        sqlRdr = Nothing
                        Exit Sub
                    End If
                End If
            End If

            If OptRecvd.Checked = True Then
                Call WareHouse_Inv_Upload()
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
            End If


        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Trans_Satus = True Then
                mP_Connection.RollbackTrans()
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
            End If
        End Try
    End Sub

    Private Sub chkDlyPullQty_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDlyPullQty.CheckStateChanged
        Try

            If chkDlyPullQty.CheckState = 1 Then
                optAvgofNextMonths.Enabled = False
                optCurMonthSch.Enabled = False
                optNextMonthSch.Enabled = False
                txtNoOfMonths.Enabled = False
            Else
                optAvgofNextMonths.Enabled = True
                optCurMonthSch.Enabled = True
                optNextMonthSch.Enabled = True
                txtNoOfMonths.Enabled = True
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CmdClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdClear.Click

        Try
            Me.spdWareHouse.MaxRows = 0
            Me.spdRelease.MaxRows = 0
            txtCustomerCode.Enabled = True
            cmdCustHelp.Enabled = True

            txtCustomerCode.Text = ""
            txtFileName.Text = ""
            txtUnitCode.Text = ""
            LblCustomerName.Text = ""
            lbltransitdaysvalue.Text = ""
            lblUnitName.Text = ""
            Me.txtConsignee.Text = ""
            optWkgDays.Checked = True
            txtDocNo.Text = ""
            FramDiff.Visible = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Try
            FramDiff.Visible = False
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Sub

    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustHelp.Click
        Try
            Dim strCustHelp() As String = Nothing
            Dim sqlWA As String = ""

            Call CmdClear_Click(CmdClear, New System.EventArgs())
            mblnDailymktUpdated = False
            mblnfilemove = False


            Dim sqlRdr2 As SqlDataReader 'Rs
            Dim strsql As String

            If OptWareHouseFile.Checked = True Then
                strCustHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code, c.cust_name from customer_mst c, " & "scheduleparameter_mst S where c.customer_code = s.customer_code and  c.UNIT_CODE = s.UNIT_CODE and c.UNIT_CODE='" & gstrUnitId & "'", "List of Customers")

            ElseIf OptReleaseFile.Checked = True Then
                strCustHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code,c.cust_name from customer_mst c," & " scheduleparameter_mst s where c.customer_code = s.customer_code and  c.UNIT_CODE = s.UNIT_CODE and c.UNIT_CODE='" & gstrUnitId & "'", " List of Customer ")
            End If

            If UBound(strCustHelp) <> -1 Then
                If strCustHelp(0) <> "0" Then
                    Me.txtCustomerCode.Text = strCustHelp(0)
                    Me.LblCustomerName.Text = strCustHelp(1)

                    If OptReleaseFile.Checked Then
                        strsql = "select top 1 ReleaseFile_Location" & " from scheduleparameter_mst" & " where customer_code = '" & txtCustomerCode.Text & "'" & " and UNIT_CODE='" & gstrUnitId & "'  order by entdt"
                        txtFileName.Text = SqlConnectionclass.ExecuteScalar(strsql)


                        strsql = "Select plant_c,plant_nm from plant_mst where UNIT_CODE='" & gstrUnitId & "'"
                        sqlRdr2 = SqlConnectionclass.ExecuteReader(strsql)

                        If sqlRdr2.HasRows Then
                            If sqlRdr2.Read() Then
                                txtUnitCode.Text = sqlRdr2.Item("plant_c")
                                lblUnitName.Text = sqlRdr2.Item("plant_nm")
                            End If
                        End If
                        sqlRdr2.Close()
                        sqlRdr2 = Nothing

                        Call CmdUploadCSV_Click(CmdUploadCSV, New System.EventArgs())
                        CmdClear.Focus()
                    End If
                Else
                    MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdDOcHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        Try

            Dim strHelp() As String
            Dim sql As Object = Nothing
            Dim sqlFileType As String = ""
            Upload_FileType = ""
            Call CmdClear_Click(CmdClear, New System.EventArgs())




            Dim sqlRdr1 As SqlDataReader
            Dim sqlRdr2 As SqlDataReader
            Dim sqlRdr3 As SqlDataReader
            Dim sqlRdr4 As SqlDataReader

            Me.txtDocNo.Text = ""
            If Upload_FileType = "EDIFACT" Then
                strHelp = ctlHelp.ShowList(gstrDSNName, gstrDatabaseName, "Select distinct convert(numeric,H1.Doc_no) as doc_no,DlyMktFlag from schedule_proposal_hdr H1, " & " SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H2 WHERE H2.DOC_NO = H1.DOC_NO and H2.UNIT_CODE = H1.UNIT_CODE and H1.UNIT_CODE= '" & gstrUnitId & "'" & " ", "Document Numbers")

            ElseIf Upload_FileType = "COVISINT" Then
                strHelp = ctlHelp.ShowList(gstrDSNName, gstrDatabaseName, "Select distinct convert(numeric,H1.Doc_no) as doc_no,DlyMktFlag from schedule_proposal_hdr H1, " & " SCHEDULE_UPLOAD_COVISINT_HDR H2 WHERE H2.DOC_NO = H1.DOC_NO and H2.UNIT_CODE = H1.UNIT_CODE and H1.UNIT_CODE= '" & gstrUnitId & "'" & " ", "Document Numbers")

            Else
                strHelp = ctlHelp.ShowList(gstrDSNName, gstrDatabaseName, "Select distinct convert(numeric,Doc_no) as doc_no,DlyMktFlag from schedule_proposal_hdr where UNIT_CODE= '" & gstrUnitId & "'", "Document Numbers")

            End If

            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    Me.txtDocNo.Text = strHelp(0)

                Else
                    MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If

                If Upload_FileType = "" Then
                    sqlFileType = "select * from schedule_upload_darwin_hdr " & " where doc_no = '" & Me.txtDocNo.Text & "' and UNIT_CODE= '" & gstrUnitId & "'"
                    sqlRdr1 = SqlConnectionclass.ExecuteReader(sqlFileType)


                    sqlFileType = " select * from schedule_upload_darwinEDIFACT_hdr" & " where doc_no = '" & Me.txtDocNo.Text & "' and UNIT_CODE= '" & gstrUnitId & "'"
                    sqlRdr1 = SqlConnectionclass.ExecuteReader(sqlFileType)


                    If sqlRdr1.HasRows Then
                        Upload_FileType = "EDIFACT"
                    End If

                    sqlFileType = " select * from schedule_upload_COVISINT_hdr" & " where doc_no = '" & Me.txtDocNo.Text & "' and UNIT_CODE= '" & gstrUnitId & "'"
                    sqlRdr1 = SqlConnectionclass.ExecuteReader(sqlFileType)

                    If sqlRdr1.FieldCount > 0 Then
                        Upload_FileType = "COVISINT"
                    End If

                End If

                If Upload_FileType = "EDIFACT" Then
                    sql = "SELECT H1.CUSTOMER_CODE,C1.CUST_NAME,H2.CONSIGNEE_CODE ,H1.StockCalcWAdays,H1.ScheduleCalcMonths," & " H2.PLANT_C,P1.PLANT_NM,H2.UPLOAD_FILE_NAME " & " FROM SCHEDULE_PROPOSAL_HDR H1,"
                    If Upload_FileType = "EDIFACT" Then
                        sql = sql & " Schedule_Upload_DarwinEDIFACT_HDR H2,"
                    End If
                    sql = sql & " CUSTOMER_MST C1, PLANT_MST P1" & " WHERE H1.DOC_NO = H2.DOC_NO AND H1.DOC_NO = '" & Me.txtDocNo.Text & "' " & " AND H1.CUSTOMER_CODE = C1.CUSTOMER_CODE AND H1.UNIT_CODE = C1.UNIT_CODE " & " AND H2.PLANT_C = P1.PLANT_C  and H2.UNIT_CODE= '" & gstrUnitId & "'"
                End If

                If Upload_FileType = "COVISINT" Then
                    sql = "SELECT H1.CUSTOMER_CODE,C1.CUST_NAME,H2.CONSIGNEE_CODE ," & " H1.StockCalcWAdays,H1.ScheduleCalcMonths, H2.plant_c ," & " p1.PLANT_NM, H2.UPLOAD_FILE_NAME " & " fROM SCHEDULE_PROPOSAL_HDR H1," & " Schedule_Upload_COVISINT_HDR H2," & " CUSTOMER_MST C1, PLANT_MST P1" & " Where H1.Doc_No = H2.Doc_No And H1.Doc_No = '" & txtDocNo.Text & "' " & " AND H1.CUSTOMER_CODE = C1.CUSTOMER_CODE " & " AND H2.PLANT_C = P1.PLANT_C  AND H2.UNIT_CODE = P1.UNIT_CODE and H2.UNIT_CODE= '" & gstrUnitId & "'"
                End If

                sqlRdr2 = SqlConnectionclass.ExecuteReader(sql)
                Me.txtCustomerCode.Text = sqlRdr2.Item("CUSTOMER_CODE")
                Me.LblCustomerName.Text = sqlRdr2.Item("CUST_NAME")
                Me.txtUnitCode.Text = sqlRdr2.Item("PLANT_C")
                Me.lblUnitName.Text = sqlRdr2.Item("PLANT_NM")
                Me.txtFileName.Text = sqlRdr2.Item("UPLOAD_FILE_NAME")
                txtConsignee.Text = sqlRdr2.Item("CONSIGNEE_CODE")

                If sqlRdr2.Item("StockCalcWAdays") = "W" Then
                    Me.optWkgDays.Checked = True
                Else
                    Me.optAvlDays.Checked = True
                End If

                If sqlRdr2.Item("ScheduleCalcMonths") = 0 Then
                    Me.optCurMonthSch.Checked = True
                Else
                    If sqlRdr2.Item("ScheduleCalcMonths") = 1 Then
                        Me.optNextMonthSch.Checked = True
                    Else
                        Me.optAvgofNextMonths.Checked = True
                    End If
                End If

                sql = "select S1.CUST_DRGNO,S1.ITEM_CODE,I1.DESCRIPTION,S1.SAFETYSTKPERDAY, " & " SAFETYDAYS,SHIPPINGQTY ,s1.shipdate," & _
                    " SPHDR.StockCalcWAdays,SPHDR.ScheduleCalcMonths " & " ,s1.consignee_code , s1.wh_code " & _
                    " from SCHEDULE_PROPOSAL_DTL S1, ITEM_MST I1 ," & " CUSTITEM_MST C,SCHEDULE_PROPOSAL_hdr SPHDR " & _
                    " WHERE C.UNIT_CODE = S1.UNIT_CODE AND C.CUST_DRGNO = S1.CUST_DRGNO AND" & _
                    " S1.ITEM_CODE = I1.ITEM_CODE and S1.UNIT_CODE = I1.UNIT_CODE " & " AND SPHDR.doc_no=S1.doc_no AND SPHDR.UNIT_CODE=S1.UNIT_CODE " & _
                    " AND S1.DOC_NO = '" & Me.txtDocNo.Text & "' S1.UNIT_CODE= '" & gstrUnitId & "'"
                Dim strsql As String
                strsql = "SELECT SHIPMENTTHRUWH FROM CUSTOMER_MST WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' and  UNIT_CODE= '" & gstrUnitId & "'"

                sqlRdr3 = SqlConnectionclass.ExecuteReader(strsql)
                If sqlRdr3.Item("SHIPMENTTHRUWH") = True Then
                    sql = sql + " and S1.CONSIGNEE_CODE = C.ACCOUNT_CODE"
                End If

                sql = sql + " order by s1.shipdate asc ,S1.CUST_DRGNO asc"


                sqlRdr2 = SqlConnectionclass.ExecuteReader(sql)
                With Me.spdRelease
                    .MaxRows = 0
                    .Row = 0

                    If sqlRdr2.FieldCount > 0 Then

                        If sqlRdr2.Item("StockCalcWAdays") = "A" Then
                            optAvlDays.Checked = True
                        ElseIf sqlRdr2.Item("StockCalcWAdays") = "W" Then
                            optWkgDays.Checked = True
                        End If
                        If sqlRdr2.Item("ScheduleCalcMonths") = 0 Then
                            optCurMonthSch.Checked = True
                        ElseIf sqlRdr2.Item("ScheduleCalcMonths") = 1 Then
                            optNextMonthSch.Checked = True
                        ElseIf sqlRdr2.Item("ScheduleCalcMonths") > 1 Then
                            optAvgofNextMonths.Checked = True
                            txtNoOfMonths.Text = sqlRdr2.Item("ScheduleCalcMonths")
                        End If
                        .Row = 1
                        While sqlRdr2.Read
                            .MaxRows = .Row
                            .SetText(Enum_Up.Cust_Drg_No, .Row, sqlRdr2.Item("CUST_DRGNO"))
                            .SetText(Enum_Up.Item_Code, .Row, sqlRdr2.Item("ITEM_CODE"))
                            .SetText(Enum_Up.Item_Desc, .Row, sqlRdr2.Item("DESCRIPTION"))
                            .SetText(Enum_Up.Sch_Qty, .Row, sqlRdr2.Item("SHIPPINGQTY").ToString)
                            .SetText(Enum_Up.SftyStkPerDay, .Row, sqlRdr2.Item("SAFETYSTKPERDAY").ToString)
                            .SetText(Enum_Up.SftyDays, .Row, sqlRdr2.Item("SAFETYDAYS").ToString)
                            .SetText(Enum_Up.Del_Dt, .Row, sqlRdr2.Item("SHIPDATE"))
                            .SetText(Enum_Up.CONSIGNEE_CODE, .Row, sqlRdr2.Item("CONSIGNEE_CODE"))
                            .SetText(Enum_Up.wh_code, .Row, sqlRdr2.Item("WH_CODE"))
                            .Row = .Row + 1
                            sqlRdr2.NextResult()

                        End While

                    End If
                End With

                sql = ""
                sql = "select top 1 dailypullflag from scheduleproposalcalculations" & " where doc_no = '" & txtDocNo.Text & "' and  UNIT_CODE= '" & gstrUnitId & "'"
                sqlRdr4 = SqlConnectionclass.ExecuteReader(sql)

                If sqlRdr4.FieldCount > 0 Then
                    If sqlRdr4.Item("DAILYPULLFLAG") = True Then
                        chkDlyPullQty.CheckState = System.Windows.Forms.CheckState.Checked
                    Else
                        chkDlyPullQty.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If
                End If

            End If


        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdFileHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFileHelp.Click
        Try

            Dim sql As String
            Dim rsPath As New ADODB.Recordset
            Dim sqlrdr As SqlDataReader

            CommanDLogOpen.FileName = ""
            CommanDLogOpen.InitialDirectory = ""
            Me.spdRelease.MaxRows = 0
            Me.spdWareHouse.MaxRows = 0
            If Me.OptReleaseFile.Checked = True Then

                sql = "SELECT RELEASEFILE_LOCATION FROM SCHEDULEPARAMETER_MST " & "WHERE CUSTOMER_CODE = '" & Me.txtCustomerCode.Text & "' and   UNIT_CODE= '" & gstrUnitId & "'"
                sqlrdr = SqlConnectionclass.ExecuteReader(sql)

                If sqlrdr.HasRows Then
                    sqlrdr.Read()
                    CommanDLogOpen.InitialDirectory = sqlrdr("ReleaseFile_Location").ToString
                Else
                    MsgBox("No Location Defined")
                    CommanDLogOpen.FileName = ""
                    CommanDLogOpen.InitialDirectory = gstrLocalCDrive
                End If
            Else
                sql = "SELECT WAREHOUSEFILE_LOCATION FROM SCHEDULEPARAMETER_MST" & _
                    " WHERE CUSTOMER_CODE = '" & Me.txtCustomerCode.Text & "' " & _
                    " AND WH_CODE = '" & Me.txtUnitCode.Text & "' and   UNIT_CODE= '" & gstrUnitId & "'"

                sqlrdr = SqlConnectionclass.ExecuteReader(sql)
                If sqlrdr.HasRows Then
                    sqlrdr.Read()
                    CommanDLogOpen.InitialDirectory = sqlrdr("WarehouseFile_Location").ToString
                Else
                    MsgBox("No Location Defined")
                    CommanDLogOpen.FileName = ""
                    CommanDLogOpen.InitialDirectory = gstrLocalCDrive
                End If
            End If

            CommanDLogOpen.Filter = "Microsoft Excel File (*.xls)|*.xls;*.CSV|Text Documents (*.Txt)|*.Txt"
            CommanDLogOpen.ShowDialog()
            Me.txtFileName.Text = CommanDLogOpen.FileName

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub Updt_DailyMkt(ByRef FileType As String)
        Dim sqlCmd As SqlCommand
        Try
            Dim intRETVAL As Integer = 0
            Dim blnOneCustDrgNo_MutipleItem As Boolean = SqlConnectionclass.ExecuteScalar("SELECT OneCustDrgNo_MutipleItem FROM customer_mst WHERE CUSTOMER_CODE='" & txtCustomerCode.Text & "' and UNIT_CODE='" & gstrUNITID & "'")

            sqlCmd = New SqlCommand
            sqlCmd.Connection = SqlConnectionclass.GetConnection
            sqlCmd.CommandType = CommandType.StoredProcedure
            sqlCmd.CommandTimeout = 0

            With sqlCmd
                If blnOneCustDrgNo_MutipleItem = False Then
                    .CommandText = "updt_dailymkt_cdp_HILEX"
                Else
                    .CommandText = "updt_dailymkt_cdp_HILEX_CustDrg_MultipleItem"
                End If
                .Parameters.Clear()
                .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 10, txtCustomerCode.Text).Direction = ParameterDirection.Input
                .Parameters.Add("@CUSTOMERCODE", SqlDbType.VarChar, 10, txtCustomerCode.Text).Direction = ParameterDirection.Input
                .Parameters.Add("@DOCNO", SqlDbType.VarChar, 10, txtDocNo.Text).Direction = ParameterDirection.Input
                .Parameters.Add("@USERID", SqlDbType.VarChar, 10, mP_User).Direction = ParameterDirection.Input
                .Parameters.Add("@FILETYPE", SqlDbType.VarChar, 10, Upload_FileType).Direction = ParameterDirection.Input
                .Parameters.Add("@RETVAL", SqlDbType.Int, 1, 0).Direction = ParameterDirection.Output

                .Parameters(0).Value = gstrUNITID
                .Parameters(1).Value = txtCustomerCode.Text
                .Parameters(2).Value = txtDocNo.Text
                .Parameters(3).Value = mP_User
                .Parameters(4).Value = Upload_FileType
                .Parameters(5).Value = ""

                .ExecuteScalar()
            End With

            Dim intCount As Integer = SqlConnectionclass.ExecuteScalar("Select count(*) from DailyMktSchedule_tempCDP  where Unit_Code='" & gstrUNITID & "' And Doc_No = " & Val(txtDocNo.Text) & "")

            If Not IsRecordExists("Select Top 1 Doc_No From dailymktschedule Where Unit_Code='" & gstrUnitId & "' And Doc_No = " & Val(txtDocNo.Text) & "") Then
                MessageBox.Show("No Schedule Data to Save.", ResolveResString(100), MessageBoxButtons.OK)
                'Added By priti on 21 Jan 2025 to add validation of drg No end date for Hilex
                If intCount > 0 Then
                    MessageBox.Show("Items are moved to Schedule distribution, Kindy use Schedule distribution", ResolveResString(100), MessageBoxButtons.OK)
                End If
                'End By priti on 17 Jan 2025 to add validation of drg No end date for Hilex
            Else
                If intCount > 0 Then
                    MessageBox.Show("Schedule Updated Successfully for Planning and Items are moved to Schedule distribution, Kindy use Schedule distribution", ResolveResString(100), MessageBoxButtons.OK)
                Else
                    MessageBox.Show("Schedule Updated Successfully for Planning.", ResolveResString(100), MessageBoxButtons.OK)
                End If
            End If

            sqlCmd.Dispose()
            sqlCmd = Nothing

        Catch ex As Exception
            If Not sqlCmd Is Nothing Then
                sqlCmd.Dispose()
                sqlCmd = Nothing
            End If
            RaiseException(ex)
        End Try
    End Sub

    Private Sub updatedailymktschedule_covisint()
        Try

            Dim STRSQL As String
            Dim sqlRdr As SqlDataReader
            Dim sqlRdr2 As SqlDataReader


            STRSQL = " SELECT DISTINCT ITEM_CODE FROM SCHEDULEPROPOSALCALCULATIONS WHERE DOC_NO='" & Me.txtDocNo.Text & "' and UNIT_CODE ='" & gstrUnitId & "'"
            sqlRdr = SqlConnectionclass.ExecuteReader(STRSQL)

            While sqlRdr.Read()
                STRSQL = ""
                STRSQL = "SELECT ITEM_CODE,MIN(SHIPMENT_DT) as SHIPMENT_DT FROM SCHEDULEPROPOSALCALCULATIONS WHERE item_code='" & sqlRdr("ITEM_CODE") & "' and DOC_NO='" & Me.txtDocNo.Text & "' and UNIT_CODE ='" & gstrUnitId & "' GROUP BY ITEM_CODE "
                sqlRdr2 = SqlConnectionclass.ExecuteReader(STRSQL)

                STRSQL = ""
                STRSQL = "UPDATE DAILYMKTSCHEDULE  SET STATUS = 0,SCHEDULE_FLAG = 0  WHERE ACCOUNT_CODE = '" & Me.txtCustomerCode.Text & "' " & " and cust_drgno= '" & sqlRdr2.Item("Item_code") & "' and UNIT_CODE ='" & gstrUnitId & "' and consignee_code in(select distinct consignee_code from Schedule_Proposal_dtl where doc_no='" & Trim(txtDocNo.Text) & "' and UNIT_CODE ='" & gstrUnitId & "') and trans_date >= '" & sqlRdr2.Item("SHIPMENT_DT") & "' and doc_no <> '" & txtDocNo.Text & "' "
                SqlConnectionclass.ExecuteNonQuery(STRSQL)

                sqlRdr.NextResult()
            End While

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdUnitHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitHelp.Click

        Try
            Dim strHelp() As String = Nothing
            Dim STRSQL As String
            Dim sqlRdr As SqlDataReader

            STRSQL = "Select distinct Customer_Mst.Cust_Name,ScheduleParameter_mst.TransitDaysbysea From ScheduleParameter_mst,Customer_Mst  Where Customer_Mst.Customer_Code=ScheduleParameter_mst.Customer_Code and Customer_Mst.UNIT_CODE=ScheduleParameter_mst.UNIT_CODE And Customer_Mst.Customer_Code = '" & Trim(Me.txtCustomerCode.Text) & "'and Customer_Mst.UNIT_CODE ='" & gstrUnitId & "'"
            sqlRdr = SqlConnectionclass.ExecuteReader(STRSQL)

            If sqlRdr.HasRows Then
                While sqlRdr.Read
                    Me.LblCustomerName.Text = sqlRdr("Cust_Name").ToString()
                    Me.lbltransitdaysvalue.Text = sqlRdr("TransitDaysBySea").ToString()
                End While
            End If
            sqlRdr.Close()

            Me.lblUnitName.Text = CStr(Nothing)
            Me.txtUnitCode.Text = CStr(Nothing)

            If OptReleaseFile.Checked = True Then
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select distinct plant_c,plant_nm from plant_mst where UNIT_CODE='" & gstrUnitId & "'")
            ElseIf OptWareHouseFile.Checked = True Then
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct C.wh_code,W.WH_DESCRIPTION  from custwarehouse_mst C,WAREHOUSE_MST W " & " where C.customer_code = '" & txtConsignee.Text & "' AND C.WH_CODE = W.WH_CODE " & " and active = 1 and  W.UNIT_CODE='" & gstrUnitId & "'")
            End If

            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    Me.txtUnitCode.Text = strHelp(0)
                    Me.lblUnitName.Text = strHelp(1)
                    If OptWareHouseFile.Checked = True Then
                        STRSQL = "select top 1 WarehouseFile_Location" & " from scheduleparameter_mst" & " where customer_code = '" & txtCustomerCode.Text & "' and WH_Code='" & Trim(txtUnitCode.Text) & "'" & " and UNIT_CODE='" & gstrUnitId & "' order by entdt"
                        sqlRdr = SqlConnectionclass.ExecuteReader(STRSQL)
                        If sqlRdr.HasRows Then
                            While sqlRdr.Read
                                txtFileName.Text = sqlRdr("WarehouseFile_Location").ToString()
                            End While
                        End If
                        sqlRdr.Close()

                    End If
                Else
                    MsgBox(" No Warehouse Defined for the selected Consignee.", MsgBoxStyle.Information, ResolveResString(100))
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Public Sub CmdUploadCSV_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdUploadCSV.Click

        Dim strsql As String = String.Empty

        Try
            Dim sourcefile As String
            If Trim(Me.txtCustomerCode.Text) = "" Then
                MsgBox("Please Select the Customer Code ", MsgBoxStyle.Information, ResolveResString(100))
                txtCustomerCode.Focus()
                Exit Sub
            End If

            If OptWareHouseFile.Checked = True Then
                If Trim(Me.txtConsignee.Text) = "" Then
                    MsgBox(" Please Select the " & lblConsignee.Text, MsgBoxStyle.Information, ResolveResString(100))
                    txtConsignee.Focus()
                    Exit Sub
                End If
            End If

            If OptWareHouseFile.Checked = True Then
                If Trim(Me.txtUnitCode.Text) = "" Then
                    MsgBox(" Please Select the " & lblUnitCode.Text, MsgBoxStyle.Information, ResolveResString(100))
                    txtUnitCode.Focus()
                    Exit Sub
                End If
            End If

            If OptWareHouseFile.Checked = True Then
                If Trim(Me.txtFileName.Text) = "" Then
                    MsgBox("Please Select the Upload File", MsgBoxStyle.Information, ResolveResString(100))
                    txtFileName.Focus()
                    Exit Sub
                End If
            End If

            If OptWareHouseFile.Checked = True Then
                Dim sqlRdr As SqlDataReader
                strsql = "select C.wh_code,W.WH_DESCRIPTION  from custwarehouse_mst C,WAREHOUSE_MST W " & " where C.customer_code = '" & txtConsignee.Text & "' AND C.WH_CODE = W.WH_CODE  AND C.UNIT_CODE = W.UNIT_CODE and c.wh_code = '" & txtUnitCode.Text & "' " & " and active = 1 and C.UNIT_CODE = '" & gstrUnitId & "'"
                sqlRdr = SqlConnectionclass.ExecuteReader(strsql)
                If Not sqlRdr.HasRows Then
                    MsgBox("Invalid Warehouse Code", MsgBoxStyle.OkOnly, ResolveResString(100))
                    txtUnitCode.Text = ""
                    sqlRdr.Close()
                    sqlRdr = Nothing
                    Exit Sub
                End If
                sqlRdr.Close()
                sqlRdr = Nothing
                FN_WareHouse_File_Upload()


            ElseIf OptReleaseFile.Checked = True Then

                Call FN_FILESELECTION()


            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Sub
    ' mayur till here
    Private Sub FN_Release_File_Upload()
        Dim Trans_Satus As Boolean
        Try


            Dim Cell_Data As String = ""
            Dim Row As Object = Nothing
            Dim i As Short = 0
            Dim Data_Row() As String = Nothing
            Dim trans_number As String = ""
            Dim Cell_Data1 As String = ""
            Dim Rev_No As Object = Nothing
            Dim Col As Short = 0
            Dim Msg As String
            Dim ShipmentFlag As Boolean
            Dim folderName As String = Nothing
            Dim strSQL As String = String.Empty
            Dim filearray(0) As Object
            Dim upldFileName(0) As Object
            Dim filedate(0) As Object
            Dim latestFile As String = Nothing

            Flag = 0

            SqlConnectionclass.ExecuteNonQuery("DELETE FROM TMPWHSTOCK where unit_code='" & gstrUnitId & "'")

            ' Mayur  101052633 
            Dim extension As String
            extension = String.Empty
            extension = Path.GetExtension(Me.txtFileName.Text)

            If extension.ToString() = ".862" Or extension.ToString() = ".830" Then
                Upload_FileType = "FORD"
                FileFord_Upload()
                Exit Sub
            End If
            ' Mayur  101052633 

            Obj_EX = New Excel.Application
            Obj_EX.Workbooks.Open(Trim(Me.txtFileName.Text))

            Row = 1
            range = Obj_EX.Cells(Row, 1)
            If Not range.Value Is Nothing Then
                Cell_Data = (range.Value.ToString)
            Else
                Cell_Data = ""
            End If

            If Len(Cell_Data) = 0 Then
                MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            Rev_No = 0

            strSQL = "SELECT iSnull(SHIPMENTTHRUWH,0) as SHIPMENTTHRUWH FROM CUSTOMER_MST" & _
            " WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' and unit_code='" & gstrUnitId & "'"
            ShipmentFlag = SqlConnectionclass.ExecuteScalar(strSQL)

            Dim dbo As New Scripting.FileSystemObject

            Trans_Satus = True
            If Len(Cell_Data) < 10 Then
                Col = 1
                range = Obj_EX.Cells(Row, Col)
                If Not range.Value.ToString Is Nothing Then
                    Cell_Data1 = (range.Value.ToString)
                Else
                    Cell_Data1 = ""
                End If
            Else
                Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
                Cell_Data1 = Trim(Data_Row(0))
            End If
            Cell_Data1 = Replace(Trim(Cell_Data1), "'", "")

            If Cell_Data1 = "DELFOR" Then
                Upload_FileType = "EDIFACT"
            ElseIf LTrim(UCase(VB.Left(Cell_Data1, 8))) = "COVISINT" Then
                Upload_FileType = "COVISINT"
            Else
                MsgBox("Wrong File Type-It's not EDIFACT/COVISINT format.", MsgBoxStyle.OkOnly, ResolveResString(100))
                Exit Sub
            End If

            If Upload_FileType = "EDIFACT" Then
                Upload_EDIFACT(Cell_Data, Row, ShipmentFlag)
            End If

            If Upload_FileType = "COVISINT" Then
                Upload_COVISINT()
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally

            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Sub
    ' till here done
    Private Sub FileFord_Upload()

        Dim Obj_FSO As New Scripting.FileSystemObject
        Dim Cell_Data As String = ""
        Dim Row As Object = Nothing
        Dim i As Short = 0
        Dim Data_Row() As String = Nothing
        Dim trans_number As String = String.Empty
        Dim fin_year_notation As String = String.Empty
        Dim Cell_Data1 As String = ""
        Dim Rev_No As Object = Nothing
        Dim Col As Short = 0
        Dim Msg As String
        Dim ShipmentFlag As Boolean
        Dim sheetNo As Short
        Dim folderName As String = Nothing
        Dim strSQL As String = String.Empty
        Dim filearray(0) As Object
        Dim upldFileName(0) As Object
        Dim filedate(0) As Object
        Dim latestFile As String = Nothing
        Dim filename As String = String.Empty
        Dim extension As String = String.Empty
        Dim file_extension As String = String.Empty
        Dim objTrans As SqlTransaction = Nothing
        Dim objConn As SqlConnection = Nothing

        Try
            extension = Path.GetFileNameWithoutExtension(Me.txtFileName.Text)
            file_extension = Path.GetExtension(Me.txtFileName.Text)
            My.Computer.FileSystem.CopyFile(Me.txtFileName.Text, "C:\CDP\CSV\" + extension + ".csv", True)

            filename = "C:\CDP\CSV\" + extension + ".csv"

            Obj_EX = New Excel.Application
            Obj_EX.Workbooks.Open(Trim(filename))

            Row = 1

            range = Obj_EX.Cells(Row, 1)
            If Not range.Value Is Nothing Then
                Cell_Data = (range.Value.ToString)
            Else
                Cell_Data = ""
            End If

            If Len(Cell_Data) = 0 Then
                MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            strSQL = "select Fin_Year_Notation from Financial_Year_Tb where unit_code = '" & gstrUnitId & "' and GETDATE() between Fin_Start_date and Fin_end_date "
            If Not IsRecordExists(strSQL) Then
                MessageBox.Show("Financial Year Not Defined.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If

            strSQL = "select Current_No from documenttype_mst where unit_code = '" & gstrUnitId & "'" & _
               " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            If Not IsRecordExists(strSQL) Then
                MessageBox.Show("Series Not Defined.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If

            objConn = New SqlConnection
            objConn = SqlConnectionclass.GetConnection()
            objTrans = objConn.BeginTransaction

            strSQL = "select Fin_Year_Notation from Financial_Year_Tb where unit_code = '" & gstrUnitId & "' and GETDATE() between Fin_Start_date and Fin_end_date "
            fin_year_notation = SqlConnectionclass.ExecuteScalar(strSQL)

            strSQL = "select Current_No from documenttype_mst where unit_code = '" & gstrUnitId & "'" & _
                " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            trans_number = SqlConnectionclass.ExecuteScalar(strSQL)

            trans_number = Val(trans_number) + 1
            strSQL = "update documenttype_mst set Current_No = " & trans_number & "  where unit_code = '" & gstrUnitId & "'" & _
               " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            SqlConnectionclass.ExecuteNonQuery(strSQL)

            trans_number = fin_year_notation + trans_number

            If Upload_ford(Cell_Data, Row, trans_number, file_extension) = False Then
                If Not objTrans Is Nothing Then
                    objTrans.Rollback()
                    objTrans = Nothing
                End If
            Else
                If Not objTrans Is Nothing Then
                    objTrans.Commit()
                    objTrans = Nothing
                End If
                Me.txtDocNo.Text = trans_number
                Call FN_FORDSCHEDULE(trans_number)
                Call Updt_DailyMkt(Upload_FileType)
                'Call MoveFile()
            End If

        Catch ex As Exception
            If Not objTrans Is Nothing Then
                objTrans.Commit()
                objTrans = Nothing
            End If
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ctlFormHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Try
            ShowHelp(("underconstruction.htm"))
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub DTPicker1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTPicker1.Validating
        Try


            Dim SQlRdr As SqlDataReader

            If txtCustomerCode.Text = "" Then
                MsgBox("Please Select the " & lblCustCode.Text, vbInformation + vbOKOnly, ResolveResString(100))
                DTPicker1.Value = GetServerDate()
                Exit Sub
            End If

            If txtConsignee.Text = "" Then
                MsgBox("Please Select the " & lblConsignee.Text, vbInformation + vbOKOnly, ResolveResString(100))
                DTPicker1.Value = GetServerDate()
                Exit Sub
            End If

            If txtUnitCode.Text = "" Then
                MsgBox("Please Select the " & lblUnitCode.Text, vbInformation + vbOKOnly, ResolveResString(100))
                DTPicker1.Value = GetServerDate()
                Exit Sub
            End If

            SQlRdr = SqlConnectionclass.ExecuteReader("SELECT MAX(TRANS_DT) as TRANS_DT FROM WAREHOUSE_STOCK_DTL" & " WHERE WAREHOUSE_CODE = '" & txtUnitCode.Text & "' " & " AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' and UNIT_CODE= '" & gstrUnitId & "'")

            If VB6.Format(DTPicker1.Value, "YYYYMMDD") < VB6.Format(SQlRdr("TRANS_DT"), "YYYYMMDD") Then
                MsgBox("You Have Already Uploaded Stock For" + vbCrLf + SQlRdr("trans_dt"), vbInformation + vbOKOnly, ResolveResString(100))
                DTPicker1.Value = GetServerDate()
            End If

            If Me.DTPicker1.Value > GetServerDate() Then
                Me.DTPicker1.Value = GetServerDate()
            End If
            SQlRdr.Close()
            SQlRdr = Nothing
        Catch ex As Exception
            RaiseException(ex)
        End Try


    End Sub


    Private Sub FRMMKTTRN0028_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            frmModules.NodeFontBold(Me.Tag) = True

            Me.DTPicker1.Value = GetServerDate()

            If OptWareHouseFile.Checked = True Then
                Me.txtConsignee.Enabled = True
                Me.cmdConsigneeHelp.Enabled = True
                Me.txtDocNo.Enabled = False
                mblnDailymktUpdated = False
                Me.optAvgofNextMonths.Checked = True
                txtNoOfMonths.Enabled = True
                txtNoOfMonths.Text = CStr(4)
                Me.optWkgDays.Checked = True
            End If

            If Me.OptReleaseFile.Checked = True Then
                Me.lblTransitDays.Visible = False
                Me.lblTransitDays.Text = "Transit Days By Sea"
                Me.lbltransitdaysvalue.Visible = False
                Me.spdRelease.Visible = True
                Me.spdWareHouse.Visible = False
            Else
                Me.lblTransitDays.Visible = False
                Me.lbltransitdaysvalue.Visible = False
                Me.spdRelease.Visible = False
                Me.spdWareHouse.Visible = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0028_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0028_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            Dim KeyCode As Short = eventArgs.KeyCode
            Dim Shift As Short = eventArgs.KeyData \ &H10000

            If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
                ctlFormHeader_ClickEvent(ctlFormHeader, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0028_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Try
            Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
            If KeyAscii = 39 Then KeyAscii = 0
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0028_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim Bool_error As Boolean = False
        Try
            bln_dateCheck = True
            Call FitToClient(Me, frmMain, ctlFormHeader, frmButton, 450)
            Call FillLabelFromResFile(Me)
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)

            Call EnableControls(True, Me)
            Me.KeyPreview = True
            Me.lblDocno.Visible = False
            Me.txtDocNo.Visible = False
            Call FN_Spread_Settings()
            Call AlignGrID()
            Me.spdRelease.Visible = False
            Me.OptWareHouseFile.Checked = True
            Me.OptStock.Checked = True
            Dim SqlRdr As SqlDataReader
            SqlRdr = SqlConnectionclass.ExecuteReader("SELECT COVISINTScheduleOverwrite from Sales_parameter where  UNIT_CODE= '" & gstrUnitId & "'")

            If SqlRdr.HasRows Then
                SqlRdr.Read()
                blnCOVISINTScheduleOverwrite = IIf(SqlRdr("COVISINTScheduleOverwrite").ToString = True, True, False)
            End If

            SqlRdr.Close()
            SqlRdr = Nothing

            If frmMKTTRN0054.bool_frm54 = True Then
                OptReleaseFile.Checked = True
                Bool_error = True
again:
                Me.txtCustomerCode.Text = cust_code.ToString
                Me.txtFileName.Text = file_name
                Call CmdUploadCSV_Click(Nothing, New System.EventArgs())
                Bool_error = False
            End If

        Catch ex As Exception
            If Bool_error = True Then
                GoTo again
            End If
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0028_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try

            Me.Dispose()
            mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FramDiff_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FramDiff.DoubleClick
        Try
            FramDiff.Visible = False
            Call Updt_DailyMkt(Upload_FileType)
            If mblnfilemove = False Then
                Call MoveFile()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub optAvgofNextMonths_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAvgofNextMonths.CheckedChanged
        If eventSender.Checked Then
            Try
                txtNoOfMonths.Text = ""
                txtNoOfMonths.Enabled = True
            Catch ex As Exception
                RaiseException(ex)
            End Try
        End If
    End Sub

    Private Sub optCurMonthSch_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCurMonthSch.CheckedChanged
        If eventSender.Checked Then
            Try
                txtNoOfMonths.Text = ""
                txtNoOfMonths.Enabled = False

            Catch ex As Exception
                RaiseException(ex)
            End Try
        End If
    End Sub

    Private Sub optNextMonthSch_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNextMonthSch.CheckedChanged
        If eventSender.Checked Then
            Try
                txtNoOfMonths.Text = ""
                txtNoOfMonths.Enabled = False

            Catch ex As Exception
                RaiseException(ex)
            End Try

        End If
    End Sub

    Private Sub OptRecvd_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptRecvd.CheckedChanged
        If eventSender.Checked Then
            Try

                spdWareHouse.MaxRows = 0
                spdWareHouse.Visible = False
                lblProposal.Visible = False

            Catch ex As Exception
                RaiseException(ex)
            End Try
        End If
    End Sub

    Private Sub OptReleaseFile_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptReleaseFile.CheckedChanged
        If eventSender.Checked Then

            Try

                Call CmdClear_Click(CmdClear, New System.EventArgs())
                chkDlyPullQty.Visible = True
                lblUnitCode.Text = "Plant Code"
                lbldate.Enabled = False
                DTPicker1.Enabled = False
                Me.spdRelease.Visible = True
                Me.spdWareHouse.Visible = False
                Me.lbltransitdaysvalue.Visible = False
                Me.lblTransitDays.Visible = False
                Me.lblTransitDays.Text = "Transit Days By Sea"
                Me.Frame3.Enabled = True
                lblProposal.Text = "Shipment Proposal Details"
                Me.txtConsignee.Enabled = False
                Me.cmdConsigneeHelp.Enabled = False

                txtUnitCode.Enabled = False
                cmdUnitHelp.Enabled = False
                txtFileName.Enabled = False
                cmdFileHelp.Enabled = False

                Me.optAvgofNextMonths.Checked = True
                txtNoOfMonths.Enabled = True
                txtNoOfMonths.Text = CStr(4)
                Me.optWkgDays.Checked = True
                Call AlignGrID()
                CmdUploadCSV.Enabled = False
                Me.lblDocno.Visible = True
                Me.txtDocNo.Visible = True
                Me.frmFileoption.Visible = False

            Catch ex As Exception
                RaiseException(ex)
            End Try
        End If
    End Sub

    Private Sub OptStock_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptStock.CheckedChanged
        If eventSender.Checked Then
            Try
                spdWareHouse.MaxRows = 0
                spdWareHouse.Visible = True
                lblProposal.Visible = True
            Catch ex As Exception
                RaiseException(ex)
            End Try
        End If
    End Sub

    Private Sub OptWareHouseFile_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptWareHouseFile.CheckedChanged
        If eventSender.Checked Then
            Try
                txtCustomerCode.Text = "" : txtUnitCode.Text = "" : txtFileName.Text = ""
                Me.spdRelease.MaxRows = 0 : Me.spdWareHouse.MaxRows = 0

                lblUnitCode.Text = "Ware House Code"
                lbldate.Enabled = True
                DTPicker1.Enabled = True

                chkDlyPullQty.Visible = False
                Me.spdRelease.Visible = False
                Me.spdWareHouse.Visible = True
                Me.lbltransitdaysvalue.Visible = False
                Me.lblTransitDays.Visible = False
                lblProposal.Text = "Warehouse Stock Details"
                Call AlignGrID()
                Me.txtConsignee.Enabled = True
                Me.cmdConsigneeHelp.Enabled = True
                Me.lblDocno.Visible = False
                Me.txtDocNo.Visible = False
                Me.frmFileoption.Visible = True

                txtUnitCode.Enabled = True
                cmdUnitHelp.Enabled = True
                txtFileName.Enabled = True
                cmdFileHelp.Enabled = True
                CmdUploadCSV.Enabled = True
                If OptRecvd.Checked = True Then
                    spdWareHouse.MaxRows = 0
                    spdWareHouse.Visible = False
                    lblProposal.Visible = False
                End If
            Catch ex As Exception
                RaiseException(ex)
            End Try
        End If
    End Sub

    Private Sub txtConsignee_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtConsignee.GotFocus
        Try
            consFocus = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtConsignee_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConsignee.KeyDown
        Try
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdConsigneeHelp_Click(cmdConsigneeHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtConsignee_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtConsignee.KeyPress
        Try
            Dim KeyAscii As Short = Asc(e.KeyChar)
            Select Case KeyAscii
                Case Keys.Enter
                    bool_Validate_Cons = False
                    txtConsignee_Validating(txtConsignee, New System.ComponentModel.CancelEventArgs((False)))
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtConsignee_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtConsignee.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Try
            consFocus = False

            If bool_Validate_Cons = True Then
                bool_Validate_Cust = False
                Exit Sub
            End If

            bool_Validate_Cust = True

            If txtCustomerCode.Text = "" Then
                MsgBox("Please Enter Customer Code", MsgBoxStyle.OkOnly, ResolveResString(100))
                Call txtCustomerCode_TextChanged(txtCustomerCode, New System.EventArgs())
                txtCustomerCode.Focus()

            End If
            If Len(Trim(Me.txtConsignee.Text)) > 0 Then
                If Not CheckRecord("Select Customer_code from Customer_mst where Customer_code = '" & Me.txtConsignee.Text & "' and UNIT_CODE= '" & gstrUnitId & "'") Then
                    MsgBox(" Invalid Consignee Code", MsgBoxStyle.Information, ResolveResString(100))
                    Me.txtConsignee.Text = "" : Me.txtConsignee.Focus()
                    Cancel = True
                End If
            End If

            If Cancel = True Then
                Me.txtConsignee.Focus()
            Else
                Me.txtUnitCode.Focus()
            End If


            eventArgs.Cancel = Cancel
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            custFocus = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            LblCustomerName.Text = ""
            lbltransitdaysvalue.Text = ""
            Me.txtUnitCode.Text = ""
            txtConsignee.Text = ""
            Me.DTPicker1.Value = GetServerDate()
            Me.lblUnitName.Text = ""
            Me.txtFileName.Text = ""
            Me.spdWareHouse.MaxRows = 0
            Me.spdRelease.MaxRows = 0
            bool_Validate_Cust = False

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            Dim KeyCode As Short = eventArgs.KeyCode
            Dim Shift As Short = eventArgs.KeyData \ &H10000
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdCustHelp_Click(cmdCustHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TxtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Try
            Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
            If KeyAscii = 13 Then
                bool_Validate_Cust = False
                TxtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs((False)))
            End If
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TxtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Try
            custFocus = False
            If bool_Validate_Cust = True Then
                bool_Validate_Cust = False
                Exit Sub
            End If


            Dim sqlRdr As SqlDataReader
            mblnDailymktUpdated = False
            mblnfilemove = False
            bool_Validate_Cust = True

            If Len(Trim(Me.txtCustomerCode.Text)) > 0 Then
                If Not CheckRecord("Select Customer_code from ScheduleParameter_mst where Customer_code = '" & Me.txtCustomerCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "'") Then
                    MsgBox(" Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                    Me.LblCustomerName.Text = "" : Me.lbltransitdaysvalue.Text = ""
                    Me.txtCustomerCode.Text = "" : Me.txtCustomerCode.Focus()
                    Cancel = True
                End If
            End If

            If Cancel = True Then
                Me.txtCustomerCode.Focus()
            Else
                sqlRdr = SqlConnectionclass.ExecuteReader("Select Customer_Mst.Cust_Name,ScheduleParameter_mst.TransitDaysbysea,customer_mst.ShipmentThruWh From ScheduleParameter_mst,Customer_Mst  Where Customer_Mst.Customer_Code=ScheduleParameter_mst.Customer_Code  and Customer_Mst.UNIT_CODE=ScheduleParameter_mst.UNIT_CODE  and ScheduleParameter_mst.UNIT_CODE= '" & gstrUnitId & "' And Customer_Mst.Customer_Code = '" & Trim(Me.txtCustomerCode.Text) & "'")
                If sqlRdr.HasRows Then
                    sqlRdr.Read()
                    Me.LblCustomerName.Text = sqlRdr("Cust_Name")
                    Me.lbltransitdaysvalue.Text = sqlRdr("TransitDaysBySea")
                    mShipmentFlag = sqlRdr("ShipmentThruWh").ToString
                End If
                Me.txtConsignee.Focus()
                If OptReleaseFile.Checked = True Then
                    txtCustomerCode.Enabled = False
                    cmdCustHelp.Enabled = False
                    Call CmdUploadCSV_Click(CmdUploadCSV, New System.EventArgs())
                    CmdClear.Focus()
                End If
                sqlRdr.Close()
                sqlRdr = Nothing

            End If


            eventArgs.Cancel = Cancel
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Function CheckRecord(ByRef strsql As String) As Boolean
        Try

            Dim sqlrdr As SqlDataReader
            sqlrdr = SqlConnectionclass.ExecuteReader(strsql)
            If sqlrdr.HasRows Then
                CheckRecord = True
            Else
                CheckRecord = False
            End If
            sqlrdr.Close()
            sqlrdr = Nothing
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Sub txtFileName_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFileName.KeyDown
        Try
            Dim KeyCode As Short = eventArgs.KeyCode
            Dim Shift As Short = eventArgs.KeyData \ &H10000
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Me.spdRelease.MaxRows = 0
                Me.spdWareHouse.MaxRows = 0
                Call cmdFileHelp_Click(cmdFileHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtNoOfMonths_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNoOfMonths.KeyPress
        Try
            Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
            If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not KeyAscii = 8 Then
                KeyAscii = 0
            End If
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtUnitCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitCode.Enter
        Try
            unitFocus = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtUnitCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitCode.GotFocus
        Try
            unitFocus = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TxtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        Try
            lblUnitName.Text = ""
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Try
            Dim KeyCode As Short = eventArgs.KeyCode
            Dim Shift As Short = eventArgs.KeyData \ &H10000
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdUnitHelp_Click(cmdUnitHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TxtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Try
            Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
            If KeyAscii = 13 Then
                Call TxtUnitCode_Validating(txtUnitCode, New System.ComponentModel.CancelEventArgs(False))
            End If
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TxtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strSQL As String = String.Empty
        Dim sqlrdr As SqlDataReader

        Try
            unitFocus = False
            If OptReleaseFile.Checked Then
                If Len(Trim(Me.txtUnitCode.Text)) > 0 Then
                    If Not CheckRecord("Select plant_c from plant_mst where plant_c = '" & Me.txtUnitCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "'") Then
                        MsgBox(" Invalid Unit Code", MsgBoxStyle.Information, ResolveResString(100))
                        Me.txtUnitCode.Text = "" : Me.txtUnitCode.Focus()
                        Cancel = True
                    End If
                End If

                If Cancel = True Then
                    Me.txtUnitCode.Focus()
                Else
                    sqlrdr = SqlConnectionclass.ExecuteReader("Select plant_nm from plant_mst where plant_c = '" & Me.txtUnitCode.Text & "'")
                    lblUnitName.Text = sqlrdr("plant_nm")
                    Me.txtFileName.Focus()
                    sqlrdr.Close()
                End If
            End If
            If OptWareHouseFile.Checked And txtUnitCode.Text <> "" Then
                sqlrdr = SqlConnectionclass.ExecuteReader("select C.wh_code,W.WH_DESCRIPTION  from custwarehouse_mst C,WAREHOUSE_MST W " & _
                                                          " where C.customer_code = '" & txtConsignee.Text & "' AND C.WH_CODE = W.WH_CODE" & _
                                                          " AND C.UNIT_CODE = W.UNIT_CODE " & " and active = 1 and C.UNIT_CODE= '" & gstrUnitId & "'")

                If sqlrdr.FieldCount = 0 Then
                    MsgBox("Invalid Warehouse Code", MsgBoxStyle.OkOnly, ResolveResString(100))
                    txtUnitCode.Text = ""
                End If
                If sqlrdr.HasRows Then
                    sqlrdr.Read()
                    strSQL = "select top 1 WarehouseFile_Location from scheduleparameter_mst" & _
                            " where customer_code = '" & txtCustomerCode.Text & "' and WH_Code='" & Trim(txtUnitCode.Text) & "'" & _
                            " and consignee_code='" & Trim(txtConsignee.Text) & "' and UNIT_CODE= '" & gstrUnitId & "' order by entdt"
                    txtFileName.Text = SqlConnectionclass.ExecuteScalar(strSQL)
                    txtFileName.Focus()
                    sqlrdr.Close()
                End If
            End If

            eventArgs.Cancel = Cancel
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub WareHouse_Inv_Upload()
        'Created By - Shubhra Verma
        'Created On - 19/Mar/2007

        Dim invobj As New prj_uploadInvoiceDaimler.prj_uploadInvoiceDaimler

        Dim RowNo As Short

        Dim SqlRdr As SqlDataReader

        Try
            Dim inv_no As Object = Nothing, ExpInv_No As Object = Nothing
            Dim Item_Suffix As String = ""
            Dim sql As String = ""
            Dim Col, Row As Short
            Dim Trans_Satus As Boolean
            Dim Stk_Qty, Item_Rate As Double

            Dim lngRevno As Integer = 0

            Dim Msg As String = ""
            Dim InvDt As String = "" ''shubhra

            If UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) = UCase("txt") Then
                Msg = invobj.WareHouse_Inv_TextFileUpload(txtCustomerCode.Text, txtUnitCode.Text, txtConsignee.Text, DTPicker1.Value, gstrConnectSQLClient, Trim(txtFileName.Text)).ToString
                MsgBox(Msg)
                Exit Sub
            End If

            sql = "select StartingRow from scheduleparameter_mst where customer_code = '" & txtCustomerCode.Text & "' and wh_code = '" & txtUnitCode.Text & "' and Consignee_code='" & Trim(txtConsignee.Text) & "' and UNIT_CODE= '" & gstrUnitId & "'"

            SqlRdr = SqlConnectionclass.ExecuteReader(sql)

            If SqlRdr.HasRows Then
                SqlRdr.Read()
                If SqlRdr("StartingRow").ToString = "" Or SqlRdr("StartingRow").ToString Is System.DBNull.Value Then
                    RowNo = 1
                Else
                    RowNo = SqlRdr("StartingRow").ToString
                End If
            End If
            SqlRdr.Close()

            Row = RowNo : Col = 1

            range = Obj_EX.Cells(Row, Col)
            If Not range.Value Is Nothing Then
                inv_no = (range.Value.ToString.Trim)
            Else
                inv_no = ""
            End If

            range = Obj_EX.Cells(Row, 3)
            If Not range.Value Is Nothing Then
                InvDt = (range.Value.ToString.Trim)
            Else
                InvDt = ""
            End If

            range = Obj_EX.Cells(Row, 2)
            If Not range.Value Is Nothing Then
                ExpInv_No = (range.Value.ToString.Trim)
            Else
                ExpInv_No = ""
            End If

            If Len(ExpInv_No) = 0 Then
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If

                MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            sql = String.Empty
            sql = "Select max(revno) as Maxrevno from InvoiceUpldWH where customer_code='" & Me.txtCustomerCode.Text & "'" & " and WareHouseCode='" & Me.txtUnitCode.Text & "' AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' and UNIT_CODE= '" & gstrUnitId & "'"
            SqlRdr = SqlConnectionclass.ExecuteReader(sql)

            If SqlRdr.HasRows Then
                SqlRdr.Read()
                lngRevno = IIf(SqlRdr("Maxrevno").ToString = "", 0, SqlRdr("Maxrevno").ToString)
            Else
                lngRevno = 0
            End If
            SqlRdr.Close()
            SqlRdr = Nothing

            lngRevno = lngRevno + 1
            mP_Connection.BeginTrans()
            mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            While Len(ExpInv_No) <> 0

                If Len(inv_no) <> 0 Then
                    sql = String.Empty
                    SqlConnectionclass.ExecuteNonQuery("insert into InvoiceUpldWH(" & "customer_code,WareHouseCode,Inv_dt,Invoice_no,revno,upld_dt,ent_dt,CONSIGNEE_CODE,UNIT_CODE)" & "values('" & Me.txtCustomerCode.Text & "'," & " '" & Me.txtUnitCode.Text & "' , " & " '" & InvDt & "' , " & " '" & inv_no & " '," & lngRevno & ",'" & DTPicker1.Value & "',getdate(),'" & txtConsignee.Text & "' ,'" & gstrUnitId & "')")
                End If

                Row = Row + 1

                range = Obj_EX.Cells(Row, Col)
                If Not range.Value Is Nothing Then
                    inv_no = (range.Value.ToString.Trim)
                Else
                    inv_no = ""
                End If

                range = Obj_EX.Cells(Row, 2)
                If Not range.Value Is Nothing Then
                    ExpInv_No = (range.Value.ToString.Trim)
                Else
                    ExpInv_No = ""
                End If

                range = Obj_EX.Cells(Row, 3)
                If Not range.Value Is Nothing Then
                    InvDt = (range.Value.ToString.Trim)
                Else
                    InvDt = ""
                End If

            End While

            sql = String.Empty
            sql = "select Invoice_no from InvoiceUpldWH " & " where REVNO= " & lngRevno & " and UNIT_CODE= '" & gstrUnitId & "' AND Invoice_no not in (" & " select CAST (doc_no AS VARCHAR) from saleschallan_dtl)" & " and warehousecode = '" & txtUnitCode.Text & "'" & " and customer_code = '" & txtCustomerCode.Text & "'" & " AND INV_DT >= '01 Jan 2008' AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' and UNIT_CODE= '" & gstrUnitId & "'"
            SqlRdr = SqlConnectionclass.ExecuteReader(sql)

            'End

            If SqlRdr.HasRows Then

                While SqlRdr.Read
                    Msg = Msg + SqlRdr("Invoice_no").ToString + " ,"
                    SqlRdr.NextResult()
                End While
                SqlRdr.Close()
                Msg = VB.Left(Msg, Len(Msg) - 1)
                MsgBox("These Invoice Nos. Are Not Defined In The System : " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                mP_Connection.RollbackTrans()
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If

                Exit Sub
            End If

            mP_Connection.CommitTrans()

            MsgBox(" Invoice Details has been Uploaded Succesfully !", MsgBoxStyle.Information, ResolveResString(100))
            'Obj_FSO = Nothing
            'If Not Obj_EX Is Nothing Then
            '    KillExcelProcess(Obj_EX)
            '    Obj_EX = Nothing
            'End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Sub

    Private Sub cmdConsigneeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConsigneeHelp.Click
        Try
            'Added By Shubhra Verma
            Dim strDocNoHelp() As String = Nothing

            If OptWareHouseFile.Checked = True Then
                If txtCustomerCode.Text = "" Then
                    MsgBox("Please Enter Customer Code.", MsgBoxStyle.OkOnly, ResolveResString(100))
                    txtConsignee.Text = ""
                    txtCustomerCode.Focus()
                    txtCustomerCode_TextChanged(txtCustomerCode, New System.EventArgs())
                    Exit Sub
                End If
                strDocNoHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code, c.cust_name from customer_mst c where c.UNIT_CODE= '" & gstrUnitId & "'", "List of Customers")

            ElseIf OptReleaseFile.Checked = True Then
                strDocNoHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code,c.cust_name from customer_mst c  where c.UNIT_CODE= '" & gstrUnitId & "'", " List of Customer ")
            End If


            If UBound(strDocNoHelp) <> -1 Then
                If strDocNoHelp(0) <> "0" Then
                    Me.txtConsignee.Text = strDocNoHelp(0)
                Else
                    MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub AlignGrID()
        Try
            If OptWareHouseFile.Checked = True Then
                Me.Frame2.Top = Me.Frame3.Top
                Me.lblProposal.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Frame2.Top) + VB6.PixelsToTwipsY(Me.Frame2.Height) + 100)
                Me.spdWareHouse.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.lblProposal.Top) + VB6.PixelsToTwipsY(Me.lblProposal.Height) + 100)
                Me.Frame3.Visible = False
            ElseIf OptReleaseFile.Checked = True Then
                Me.Frame3.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Frame1.Top) + VB6.PixelsToTwipsY(Me.Frame1.Height) + 100)
                Me.Frame2.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Frame3.Top) + VB6.PixelsToTwipsY(Me.Frame3.Height) + 50)
                Me.lblProposal.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Frame2.Top) + VB6.PixelsToTwipsY(Me.Frame2.Height) + 50)
                Me.spdWareHouse.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.lblProposal.Top) + VB6.PixelsToTwipsY(Me.lblProposal.Height) + 50)
                Me.Frame3.Visible = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Function Upload_COVISINT() As Boolean

        Dim Flag As Short
        Dim SqlRdr As SqlDataReader
        Dim Cell_Data As String
        Dim Row, i As Short
        Dim Data_Row() As String
        Dim Trans_Satus As Boolean
        Dim Upload_FileType, sql, Cell_Data1 As String
        Dim WA As String
        Dim Sch As Integer = 0
        Dim HOLIDAY As String
        Dim Msg As String
        Dim Consignee_Code As String
        Dim sqlWarehouse As String
        Dim SheetCount As Integer
        Dim CustDrgNo As Object = Nothing
        Dim ItemCode As Object = Nothing
        Dim SftyStk As Object = Nothing
        Dim SftyDays As Object = Nothing
        Dim ShpgQty As Object = Nothing
        Dim ShipDate As Object = Nothing
        Dim wh_code As Object = Nothing
        Dim trans_number As String
        Dim fin_year_notation As String
        Dim strSQL As String = String.Empty
        Dim objConn As SqlConnection = Nothing
        Dim objTrans As SqlTransaction = Nothing

        Try

            Obj_EX = New Excel.Application
            Obj_EX.Workbooks.Open(Trim(txtFileName.Text))

            Row = 1

            range = Obj_EX.Cells(Row, 1)
            If Not range.Value Is Nothing Then
                Cell_Data = (range.Value.ToString)
            Else
                Cell_Data = ""
            End If

            If Len(Cell_Data) = 0 Then
                MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            If mShipmentFlag = True Then
                Dim doc_code As String
                Dim Cust_Vendor_Code As String

                range = Obj_EX.Cells(3, 2)
                If Not range.Value.ToString Is Nothing Then
                    doc_code = (range.Value.ToString)
                Else
                    doc_code = ""
                End If

                range = Obj_EX.Cells(2, 2)
                If Not range.Value.ToString Is Nothing Then
                    Cust_Vendor_Code = (range.Value.ToString)
                Else
                    Cust_Vendor_Code = ""
                End If

                If doc_code = "" Or Cust_Vendor_Code = "" Then
                    MsgBox("Shipment for this Customer is through Warehouse" + vbCrLf + "But Dock Code or Cust Vend Code Not Defined" + vbCrLf + "in the Release File" + vbCrLf + "Please Check...", MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
            End If

            strSQL = "select Fin_Year_Notation from Financial_Year_Tb where unit_code = '" & gstrUnitId & "' and GETDATE() between Fin_Start_date and Fin_end_date "
            If Not IsRecordExists(strSQL) Then
                MessageBox.Show("Financial Year Not Defined.", ResolveResString(100), MessageBoxButtons.OK)
                Return False
            End If


            strSQL = "select Current_No from documenttype_mst where unit_code = '" & gstrUnitId & "'" & _
               " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            If Not IsRecordExists(strSQL) Then
                MessageBox.Show("Series Not Defined.", ResolveResString(100), MessageBoxButtons.OK)
                Return False
            End If

            objConn = New SqlConnection
            objConn = SqlConnectionclass.GetConnection()
            objTrans = objConn.BeginTransaction
            Trans_Satus = True

            strSQL = "select Fin_Year_Notation from Financial_Year_Tb where unit_code = '" & gstrUnitId & "' and GETDATE() between Fin_Start_date and Fin_end_date "
            fin_year_notation = SqlConnectionclass.ExecuteScalar(strSQL)

            strSQL = "select Current_No from documenttype_mst where unit_code = '" & gstrUnitId & "'" & _
                " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            trans_number = SqlConnectionclass.ExecuteScalar(strSQL)

            trans_number = Val(trans_number) + 1
            strSQL = "update documenttype_mst set Current_No = " & trans_number & "  where unit_code = '" & gstrUnitId & "'" & _
               " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            SqlConnectionclass.ExecuteNonQuery(strSQL)

            trans_number = fin_year_notation + trans_number

            sql = " Insert Into schedule_upload_covisint_hdr(doc_no,doc_type,cust_code," & _
                " plant_c,upload_file_name,upload_file_type,ent_dt,ent_uid,updt_dt,updt_uid,DOC_DT," & _
                " CONSIGNEE_CODE,UNIT_CODE) " & " Values (" & trans_number & ",302,'" & Trim(txtCustomerCode.Text) & "'" & _
                " ," & " '" & txtUnitCode.Text & "', " & " '" & txtFileName.Text.Trim & "'," & _
                " ''," & " getDate(),'" & mP_User & "' ," & " getDate()," & _
                " '" & mP_User & "',getDate(),'" & txtConsignee.Text.Trim & "','" & gstrUnitId & "') "

            SqlConnectionclass.ExecuteNonQuery(sql)

            SheetCount = 1
            While SheetCount <= Obj_EX.Sheets.Count
                HOLIDAY = ""
                Obj_EX.Sheets.Item(SheetCount).Select()
                Row = 1
                range = Obj_EX.Cells(Row, 1)
                If Not range.Value.ToString Is Nothing Then
                    Cell_Data = (range.Value.ToString)
                Else
                    Cell_Data = ""
                End If

                If Len(Cell_Data) = 0 Then
                    MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
                    If Not objTrans Is Nothing Then
                        objTrans.Rollback()
                        objTrans = Nothing
                    End If
                    Return False
                End If

                Upload_FileType = "R"

                If mShipmentFlag = True Then
                    Dim doc_code As String
                    Dim Cust_Vendor_Code As String

                    range = Obj_EX.Cells(3, 2)
                    If Not range.Value Is Nothing Then
                        doc_code = (range.Value.ToString)
                    Else
                        doc_code = ""
                    End If

                    range = Obj_EX.Cells(2, 2)
                    If Not range.Value Is Nothing Then
                        Cust_Vendor_Code = (range.Value.ToString)
                    Else
                        Cust_Vendor_Code = ""
                    End If

                    sql = "select customer_code from customer_mst " & "where dock_code = '" & Trim(doc_code) & "'" & _
                        "and Cust_Vendor_Code = '" & Trim(Cust_Vendor_Code) & "'  and  UNIT_CODE= '" & gstrUnitId & "'"
                    SqlRdr = SqlConnectionclass.ExecuteReader(sql)

                    Consignee_Code = SqlRdr("customer_code").ToString()
                    SqlRdr.Close()
                Else
                    Consignee_Code = txtCustomerCode.Text
                End If

                Row = 6
                i = 2
                Dim trnNo1 As String
                Dim trnNo2 As String
                Dim trnNo3 As String
                Dim trnNo4 As String
                Dim trnNo5 As String

range1:
                range = Obj_EX.Cells(5, i)
                While Not range.Value Is Nothing
                    If Len(Trim(range.Value)) = 0 Then
                        If Not objTrans Is Nothing Then
                            objTrans.Rollback()
                            objTrans = Nothing
                        End If
                        Return False
                    End If
range2:
                    range = Obj_EX.Cells(Row, i)
                    While Not range.Value Is Nothing
                        If Len(Trim(range.Value)) = 0 Then
                            If Not objTrans Is Nothing Then
                                objTrans.Rollback()
                                objTrans = Nothing
                            End If
                            Return False
                        End If
                        range = Obj_EX.Cells(5, i)

                        If Not range.Value Is Nothing Then
                            trnNo1 = range.Value.ToString
                        Else
                            trnNo1 = ""
                        End If

                        range = Obj_EX.Cells(Row, i)
                        If Not range.Value Is Nothing Then
                            trnNo2 = range.Value.ToString
                        Else
                            trnNo2 = ""
                        End If

                        range = Obj_EX.Cells(2, 2)
                        If Not range.Value Is Nothing Then
                            trnNo3 = range.Value.ToString
                        Else
                            trnNo3 = ""
                        End If

                        range = Obj_EX.Cells(3, 2)
                        If Not range.Value Is Nothing Then
                            trnNo4 = range.Value.ToString
                        Else
                            trnNo4 = ""
                        End If

                        sql = "Insert into schedule_upload_covisint_dtl(doc_no,doc_type, " & _
                            " item_code,WH_CODE,factory_code,consignee_code,delivery_date,qty,ent_dt,ent_uid,updt_dt,updt_uid,UNIT_CODE)" & _
                            " Values (" & trans_number & ",302,'" & Trim(trnNo1) & "'," & " '" & Trim(trnNo3) & "','" & Trim(trnNo4) & "'," & _
                            " '" & Trim(Consignee_Code) & "',Convert(DateTime, '" & Trim(IIf(Obj_EX.Range("$A$" & Row).Value Is Nothing, "01/01/1900", Obj_EX.Range("$A$" & Row).Value.ToString)) & "', 103) ," & _
                            " '" & Trim(trnNo2) & "',getDate()," & " '" & mP_User & "',getDate() ,'" & mP_User & "','" & gstrUnitId & "')"
                        SqlConnectionclass.ExecuteNonQuery(sql)

                        sql = "select Distinct dt from Calendar_mkt_Cust " & _
                            " where dt = Convert(DateTime, '" & Trim(IIf(Obj_EX.Range("$A$" & Row).Value Is Nothing, "01/01/1900", Obj_EX.Range("$A$" & Row).Value.ToString)) & "', 103) " & _
                            " AND work_flg = 1 and UNIT_CODE= '" & gstrUnitId & "' and Cust_Code = (SELECT CUSTOMER_CODE FROM " & _
                            " CUSTOMER_MST WHERE CUST_VENDOR_CODE = '" & Trim(Consignee_Code) & "' and UNIT_CODE= '" & gstrUnitId & "')"

                        SqlRdr = SqlConnectionclass.ExecuteReader(sql)

                        If SqlRdr.HasRows Then
                            While SqlRdr.Read
                                If InStr(Replace(HOLIDAY, SqlRdr("dt").ToString, "$"), "$") = 0 Then
                                    HOLIDAY = HOLIDAY & " " & SqlRdr("dt").ToString  'Replace used By Amit
                                End If
                                SqlRdr.Close()
                            End While
                        End If
                        Row = Row + 1
                        GoTo range2
                    End While

                    i = i + 1
                    Row = 6
                    GoTo range1
                End While

                HOLIDAY = Replace(HOLIDAY, " ", Chr(13))
                If HOLIDAY <> "" Then
                    If Not objTrans Is Nothing Then
                        objTrans.Rollback()
                        objTrans = Nothing
                    End If
                    MsgBox(HOLIDAY & vbCrLf & "  is/are not working day(s) ")
                    Flag = 1
                End If
                SheetCount = SheetCount + 1
            End While

            Dim countREC As Short

            sql = "select DISTINCT C.ACCOUNT_CODE, C.cust_drgno,COUNT(C.item_code) countitem from custitem_mst C with (nolock)" & _
                " where C.active = 1 AND SCHUPLDREQD = 1 and  unit_code='" & gstrUnitId & "' and "

            If mShipmentFlag = True Then
                sql = sql & " C.ACCOUNT_CODE IN (SELECT DISTINCT CONSIGNEE_CODE"
                sql = sql & " FROM SCHEDULE_UPLOAD_COVISINT_DTL with (nolock) where UNIT_CODE ='" & gstrUnitId & "' )" & _
                        " AND C.CUST_DRGNO IN (SELECT DISTINCT ITEM_CODE"
            Else
                sql = sql & " C.ACCOUNT_CODE = '" & txtCustomerCode.Text & "'"
                sql = sql & " AND C.CUST_DRGNO IN (SELECT DISTINCT ITEM_CODE"
            End If

            sql = sql & " FROM SCHEDULE_UPLOAD_COVISINT_DTL with (nolock)"
            sql = sql & " WHERE DOC_NO = '" & trans_number & "'  and  UNIT_CODE ='" & gstrUnitId & "') " & " group by C.Account_code, C.cust_drgno"

            SqlRdr = SqlConnectionclass.ExecuteReader(sql)

            Msg = ""
            '' Added by priti 
            Dim blnOneCustDrgNo_MutipleItem As Boolean = SqlConnectionclass.ExecuteScalar("SELECT OneCustDrgNo_MutipleItem FROM customer_mst WHERE CUSTOMER_CODE='" & txtCustomerCode.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
            If SqlRdr.HasRows Then
                While SqlRdr.Read
                    countREC = SqlRdr("COUNTITEM").ToString
                    If countREC > 1 Then
                        Dim strCustDrgNo = SqlRdr("cust_drgno").ToString
                        strSQL = "Update SCHEDULE_UPLOAD_COVISINT_DTL set IsOneCustDrgNo_MutipleItem=1 WHERE DOC_NO = '" & trans_number & "'  and  UNIT_CODE ='" & gstrUNITID & "' and Item_code='" & strCustDrgNo & "'"
                        SqlConnectionclass.ExecuteNonQuery(strSQL)
                        If blnOneCustDrgNo_MutipleItem = False Then
                            Msg = Msg & "  " + SqlRdr("cust_drgno").ToString
                            Flag = 1
                        End If
                    End If
                End While
                If Msg <> "" Then
                    If Not objTrans Is Nothing Then
                        objTrans.Rollback()
                        objTrans = Nothing
                    End If
                    MsgBox("Following Cust_DrgNo(s) Are Active For Multiple Items " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    Return False
                End If
            End If

            SqlRdr.Close()
            SqlRdr = Nothing

            sql = "select distinct d.ITEM_CODE " & " from Schedule_Upload_COVISINT_Dtl d,Schedule_Upload_COVISINT_hdr h " & " Where  h.cust_code = '" & Me.txtCustomerCode.Text & "'" & "and  d.UNIT_CODE ='" & gstrUnitId & "' and d.doc_no = h.doc_no and d.UNIT_CODE = h.UNIT_CODE and h.doc_no=" & trans_number & "" & " and ltrim(rtrim(d.ITEM_CODE)) " & " not in (select cust_drgno " & " from custitem_mst where active = 1 AND SCHUPLDREQD = 1 and  UNIT_CODE ='" & gstrUnitId & "'"

            If mShipmentFlag = True Then
                sql = sql & " and account_code in (SELECT DISTINCT CONSIGNEE_CODE FROM SCHEDULE_UPLOAD_COVISINT_DTL WHERE DOC_NO = " & trans_number & " and   UNIT_CODE ='" & gstrUnitId & "'))"
            Else
                sql = sql & " and account_code = '" & Me.txtCustomerCode.Text & "')"
            End If

            sql = sql & " UNION "
            sql = sql & " select distinct d.ITEM_CODE " & " from Schedule_Upload_COVISINT_Dtl d,Schedule_Upload_COVISINT_hdr h " & _
                " Where  h.cust_code = '" & Me.txtCustomerCode.Text & "'" & _
                "and  d.UNIT_CODE ='" & gstrUnitId & "' and d.UNIT_CODE = h.UNIT_CODE  and d.doc_no = h.doc_no and h.doc_no=" & trans_number & "" & _
                " and ltrim(rtrim(d.ITEM_CODE)) " & _
                " not in (select cust_drgno " & _
                " from custitem_mst where active = 1 AND SCHUPLDREQD = 1 and  UNIT_CODE ='" & gstrUnitId & "'"
            sql = sql & " and account_code = '" & Me.txtCustomerCode.Text & "')"

            SqlRdr = SqlConnectionclass.ExecuteReader(sql)

            Msg = ""
            If SqlRdr.HasRows Then
                While SqlRdr.Read
                    Msg = Msg & "'" + SqlRdr("ITEM_CODE").ToString + "'" + vbCrLf
                End While

                If Len(Trim(Msg)) > 0 Then
                    If Not objTrans Is Nothing Then
                        objTrans.Rollback()
                        objTrans = Nothing
                    End If
                    MsgBox("Following Items Are Not Defined In The System" & vbCrLf & _
                           "for Customer " & txtCustomerCode.Text & " : " & vbCrLf & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Return False
                End If

            End If

            SqlRdr.Close()
            SqlRdr = Nothing

            HOLIDAY = Replace(HOLIDAY, " ", Chr(13))
            If HOLIDAY <> "" Then
                If Not objTrans Is Nothing Then
                    objTrans.Rollback()
                    objTrans = Nothing
                End If
                MsgBox("Following is/are not working day(s):" & vbCrLf & vbCrLf & "Consignee---Date" & vbCrLf & HOLIDAY & vbCrLf)
            End If

            sql = "select distinct d.WH_CODE from schedule_upload_COVISINT_hdr H," & _
                    " schedule_upload_COVISINT_dtl D,scheduleparameter_mst s " & _
                    " where h.unit_code = d.unit_code and h.doc_no = d.doc_no and d.unit_code = s.unit_code and" & _
                    " h.cust_code = s.customer_code and " & _
                    " D.WH_CODE  not in(select D.wh_code " & _
                    " from scheduleparameter_mst s where s.UNIT_CODE= '" & gstrUnitId & "'" & _
                    " and   s.customer_code =  '" & Me.txtCustomerCode.Text & "')" & _
                    " and H.cust_code = '" & Me.txtCustomerCode.Text & "' and D.doc_no = " & trans_number & " and H.UNIT_CODE= '" & gstrUnitId & "'  "
            SqlRdr = SqlConnectionclass.ExecuteReader(sql)

            Msg = String.Empty

            If mShipmentFlag = True Then
                If SqlRdr.HasRows Then
                    SqlRdr.Read()
                    Msg = Msg & "  '" + SqlRdr("WH_Code").ToString() + "'  "
                    Flag = 1
                End If

                If Msg <> "" Then
                    If Not objTrans Is Nothing Then
                        objTrans.Rollback()
                        objTrans = Nothing
                    End If
                    MsgBox("WRONG WAREHOUSE: " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    Return False
                End If
            End If
            SqlRdr.Close()
            SqlRdr = Nothing

            If Me.optWkgDays.Checked = True Then
                WA = "W"
            Else
                WA = "A"
            End If

            If Me.optCurMonthSch.Checked = True Then
                Sch = 0
            ElseIf Me.optNextMonthSch.Checked = True Then
                Sch = 1
            Else
                Sch = Val(Me.txtNoOfMonths.Text)
            End If

            If Flag = 1 Then
                Return False
            Else
                If mShipmentFlag = False And chkDlyPullQty.Checked = True Then
                    If Not objTrans Is Nothing Then
                        objTrans.Rollback()
                        objTrans = Nothing
                    End If
                    MsgBox("Shipment For This Customer Is Not Through Warehouse, " & vbCrLf & "So, You Can't Use Daily Pull Qty For Safety Stock Calculations.", vbInformation, ResolveResString(100))
                    Return False
                Else
                    If Not objTrans Is Nothing Then
                        objTrans.Commit()
                        objTrans = Nothing
                    End If
                    Me.txtDocNo.Text = trans_number
                End If
            End If

            Dim intcountWH As Integer

            If mShipmentFlag = True Then
                SqlConnectionclass.ExecuteNonQuery("EXEC SP_CALCULATESAFETYSTOCKFORSCHEDULE_COVISINT '" & gstrUnitId & "','" & Me.txtCustomerCode.Text & "'," & " '" & trans_number & "','" & WA & "','" & Sch & "','" & gstrIpaddressWinSck & "'")

                sql = "Select count(distinct WH_CODE) count from SCHEDULE_UPLOAD_COVISINT_DTL with (nolock) where doc_no ='" & trans_number & "' and UNIT_CODE= '" & gstrUnitId & "' group by WH_CODE"
                intcountWH = SqlConnectionclass.ExecuteScalar(sql)

                sql = "Select WH_CODE from SCHEDULE_UPLOAD_COVISINT_DTL with (nolock)  where WH_CODE not in" & _
                    " (Select distinct WareHouse_Code from WareHouse_Stock_Dtl where  UNIT_CODE= '" & gstrUnitId & "')" & _
                    " and UNIT_CODE= '" & gstrUnitId & "'"

                If IsRecordExists(sql) Then
                    MsgBox("Stock is not defined for the Warehouse(s).So no Schedule will be proposed.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Return False
                End If
                If intcountWH > 1 Then
                    MsgBox("You have the Release File with more than 1 warehouses." & vbCrLf & "Details for these will be available in Schedule Proposal Details.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
                Call FN_Display(trans_number, Upload_FileType)
            Else
                Call FN_Display_WITHOUTWH(trans_number)
            End If

            strSQL = "Select Top 1 Doc_No From ScheduleProposalcalculations Where Doc_No= '" & Val(trans_number) & "' and shipment_qty > 0 AND UNIT_CODE = '" & gstrUnitId & "'"

            If Not IsRecordExists(strSQL) Then
                MessageBox.Show("No Schedule Proposed.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Function
            Else
                MessageBox.Show("Schedule has been Uploaded Succesfully.", ResolveResString(100), MessageBoxButtons.OK)
                objConn = Nothing
                objConn = New SqlConnection
                objConn = SqlConnectionclass.GetConnection()
                objTrans = objConn.BeginTransaction
                Call Updt_DailyMkt(Upload_FileType)
                objTrans.Commit()
                objTrans = Nothing
                objConn = Nothing
                AutoMailerSend(Val(trans_number))
            End If

            'With Me.spdRelease
            '    If .MaxRows > 0 Then

            '        SqlConnectionclass.ExecuteNonQuery("Insert into Schedule_proposal_HDR(doc_no,doc_dt,customer_code,StockCalcWAdays,ScheduleCalcMonths,DlyMktFlag,UNIT_CODE ) values('" & trans_number & "','" & Convert.ToDateTime(Me.DTPicker1.Value).ToString("dd MMM yyyy") & "', '" & Me.txtCustomerCode.Text & "' ,'" & WA & "'," & Sch & ",0,'" & gstrUnitId & "')")
            '        .Row = 1
            '        While .Row <= .MaxRows
            '            CustDrgNo = Nothing
            '            ItemCode = Nothing
            '            SftyStk = Nothing
            '            SftyDays = Nothing
            '            ShpgQty = Nothing
            '            ShipDate = Nothing
            '            wh_code = Nothing
            '            Consignee_Code = Nothing

            '            Call .GetText(Enum_Up.Cust_Drg_No, .Row, CustDrgNo)
            '            Call .GetText(Enum_Up.Item_Code, .Row, ItemCode)
            '            Call .GetText(Enum_Up.SftyStkPerDay, .Row, SftyStk)
            '            Call .GetText(Enum_Up.SftyDays, .Row, SftyDays)
            '            Call .GetText(Enum_Up.Sch_Qty, .Row, ShpgQty)
            '            Call .GetText(Enum_Up.Del_Dt, .Row, ShipDate)
            '            Call .GetText(Enum_Up.wh_code, .Row, wh_code)
            '            Call .GetText(Enum_Up.CONSIGNEE_CODE, .Row, Consignee_Code)
            '            If SftyStk = "" Then
            '                SftyStk = 0
            '            End If

            '            If SftyDays = "" Then
            '                SftyDays = 0
            '            End If

            '            If ShpgQty.ToString = "" Then
            '                ShpgQty = 0
            '            End If

            '            SqlConnectionclass.ExecuteNonQuery(" insert into Schedule_Proposal_Dtl(Doc_no,Cust_DrgNo,Item_code,safetyStkPerDay,safetyDays," & _
            '                    " shippingQty,shipdate,WH_Code,CONSIGNEE_CODE,UNIT_CODE) values( " & trans_number & "," & " '" & CustDrgNo & "','" & ItemCode & "', " & SftyStk & " ," & " " & SftyDays & ", " & ShpgQty & ",'" & Convert.ToDateTime(ShipDate).ToString("dd MMM yyyy") & "','" & wh_code & "','" & Consignee_Code & "' ,'" & gstrUnitId & "') ")

            '            .Row = .Row + 1

            '        End While

            '    End If
            'End With

            strSQL = "Select Top 1 Doc_No From dailymktschedule Where Unit_Code='" & gstrUnitId & "'" & _
                                            " and Doc_No=" & Val(trans_number)

            Dim YesNo As String

            If Not IsRecordExists(strSQL) Then
                YesNo = CStr(MsgBox("No Schedule Proposed." + vbCrLf + "Do You Want To Move Files To BackUp Folder?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)))
                If YesNo = CStr(MsgBoxResult.Yes) Then Call MoveFile()
                If YesNo = CStr(MsgBoxResult.No) Then
                    Exit Function
                End If
            Else
                MsgBox(" Schedule has been Uploaded Succesfully !", MsgBoxStyle.Information, ResolveResString(100))
                'Call FillReleaseDiffDetails()
                'If FramDiff.Visible = False Then
                '    If mblnDailymktUpdated = False Then
                '        Call Updt_DailyMkt(Upload_FileType)
                If mblnfilemove = False Then
                    Call MoveFile()
                End If
                'End If
                '    End If
            End If

            If Not SqlRdr Is Nothing Then
                SqlRdr.Close()
                SqlRdr = Nothing
            End If

            Return True

        Catch ex As Exception
            If Not objTrans Is Nothing Then
                objTrans.Rollback()
                objTrans = Nothing
            End If
            RaiseException(ex)
            Return False
        Finally
            If SqlRdr.IsClosed = False Then
                SqlRdr.Close() : SqlRdr = Nothing
            End If

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

        End Try
    End Function

    Private Function FN_Display_WITHOUTWH(ByVal TRANS_NUMBER As String) As Object
        'ISSUE ID - eMpro-20080410-17008
        Dim Transit_Days, Row, Buffer_Days As Integer
        Dim CustDrgNo As Object = Nothing, DELDT As Object = Nothing
        Dim SCHQTY As Double
        Dim SFTYDAYS_MNTD As Object = Nothing
        Dim sql As String = "", updSQL As String = "", strWHCode As String = ""
        Dim lngBagQty As Long
        Dim objDate As Object = Nothing
        Dim sqlRdr As SqlDataReader
        Dim objConn As SqlConnection = Nothing
        Dim objTrans As SqlTransaction = Nothing

        Try
            If Upload_FileType = "EDIFACT" Or Upload_FileType = "COVISINT" Then

                If Upload_FileType = "EDIFACT" Then

                    sql = "Select Distinct D.Delivery_Dt,h.PARTY_ID1 ,H.PARTY_ID3," & _
                          " C.Cust_Drgno,I.ITEM_Code,I.DESCRIPTION, " & _
                           "QUANTITY AS SHIPQTY,Frequency,DISPATCH_PATTERN,D.RAN_NO"


                    sql = sql & " From SCHEDULE_UPLOAD_DARWINEDIFACT_DTL D, CUSTITEM_MST C,ITEM_MST I,  " & _
                        "SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H, SCHEDULEPARAMETER_mst SP"
                    sql = sql & " Where C.Account_code=SP.Customer_code and C.Active=1"
                    sql = sql & " And H.Doc_No=" & TRANS_NUMBER & ""

                    sql = sql & " And C.UNIT_CODE=SP.UNIT_CODE"
                    sql = sql & " And D.UNIT_CODE=C.UNIT_CODE"
                    sql = sql & " And I.UNIT_CODE=C.UNIT_CODE"
                    sql = sql & " And D.UNIT_CODE=H.UNIT_CODE"
                    sql = sql & " And D.UNIT_CODE='" & gstrUnitId & "'"

                    sql = sql & " AND D.ITEM_CODE = C.CUST_DRGNO AND C.ITEM_CODE = I.ITEM_CODE"
                    sql = sql & " AND D.DOC_NO = H.DOC_NO and ltrim(rtrim(frequency))<>'' and ltrim(rtrim(dispatch_pattern))<>'' " & _
                                " AND SP.CUSTOMER_CODE  = '" & Me.txtCustomerCode.Text & "' "
                    sql = sql & " Order By D.DELIVERY_DT "
                End If

                If Upload_FileType = "COVISINT" Then

                    sql = "Select Distinct D.Delivery_DATE AS DELIVERY_DT,D.WH_CODE AS PARTY_ID1,C.Cust_Drgno,I.ITEM_Code," & _
                      " I.DESCRIPTION, QTY AS SHIPQTY,FACTORY_CODE " & _
                      " From SCHEDULE_UPLOAD_COVISINT_DTL D, CUSTITEM_MST C,ITEM_MST I," & _
                      " SCHEDULE_UPLOAD_COVISINT_HDR H, SCHEDULEPARAMETER_mst SP" & _
                      " Where C.Account_code=SP.Customer_code and" & _
                      " C.Active=1 And H.Doc_No = " & TRANS_NUMBER & "  AND D.ITEM_CODE = C.CUST_DRGNO" & _
                      " AND D.UNIT_CODE = C.UNIT_CODE AND C.UNIT_CODE = I.UNIT_CODE AND" & _
                      " H.UNIT_CODE = I.UNIT_CODE AND H.UNIT_CODE = SP.UNIT_CODE AND H.UNIT_CODE = '" & gstrUnitId & "'" & _
                      " AND C.ITEM_CODE = I.ITEM_CODE AND D.DOC_NO = H.DOC_NO" & _
                      " and SP.CUSTOMER_CODE  = '" & txtCustomerCode.Text & "'  Order By D.DELIVERY_DATE"
                End If

                sqlRdr = SqlConnectionclass.ExecuteReader(sql)

                SCHQTY = 0

                'Row = 0 : spdRelease.MaxRows = Row

                If Not sqlRdr.HasRows Then Exit Function

                sql = " Select isnull(TransitDaysBySea,0) From ScheduleParameter_mst" & _
                           " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                Transit_Days = SqlConnectionclass.ExecuteScalar(sql)

                sql = " Select isnull(BufferDays,0) From ScheduleParameter_mst" & _
                        " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                Buffer_Days = SqlConnectionclass.ExecuteScalar(sql)

                objConn = New SqlConnection
                objConn = SqlConnectionclass.GetConnection()
                objTrans = objConn.BeginTransaction

                While sqlRdr.Read
                    sql = "select isnull(bag_qty,1) from item_mst where item_code = '" & sqlRdr("Item_Code").ToString & "'" & _
                        " and UNIT_CODE = '" & gstrUnitId & "'"
                    mlngBAGQTY = SqlConnectionclass.ExecuteScalar(sql)
                    If mlngBAGQTY = 0 Then mlngBAGQTY = 1
                    mlngBAGQTY = 1
                    sql = " select max(dt) as dt from Calendar_Mfg_mst" & _
                        " where work_flg=0 and dt < = '" & getDateForDB(DateAdd("d", -(Transit_Days + Buffer_Days), sqlRdr("DELIVERY_DT").ToString)) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                    objDate = SqlConnectionclass.ExecuteScalar(sql)

                    If Upload_FileType = "EDIFACT" Then
                        If sqlRdr("Frequency").ToString = "" And sqlRdr("DISPATCH_PATTERN").ToString = "" Then
                            GoTo SKIP
                        End If
                    End If

                    If DateDiff(Microsoft.VisualBasic.DateInterval.Day, objDate, ServerDate()) > 0 Or IIf(IsDBNull(sqlRdr("shipqty").ToString), 0, sqlRdr("shipqty").ToString) = 0 Then
                        GoTo SKIP
                    End If


                    SCHQTY = SCHQTY + sqlRdr("shipqty").ToString
                    'Row = Row + 1 : Me.spdRelease.MaxRows = Row

                    If Val(mlngBAGQTY) >= Val(SCHQTY) Then
                        SCHQTY = mlngBAGQTY
                    Else
                        If mlngBAGQTY > 0 Then
                            If SCHQTY Mod mlngBAGQTY > 0 Then
                                SCHQTY = (mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)) + mlngBAGQTY
                            Else
                                SCHQTY = mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)
                            End If
                        End If
                    End If

                    'Me.spdRelease.SetText(Enum_Up.Del_Dt, Row, IIf(IsDBNull(sqlRdr("DELIVERY_DT").ToString), "", getDateForDB(sqlRdr("DELIVERY_DT").ToString)))
                    'Me.spdRelease.SetText(Enum_Up.Cust_Drg_No, Row, IIf(IsDBNull(sqlRdr("Cust_DrgNo").ToString), "", sqlRdr("Cust_DrgNo").ToString))
                    'Me.spdRelease.SetText(Enum_Up.Item_Code, Row, IIf(IsDBNull(sqlRdr("Item_Code").ToString), "", sqlRdr("Item_Code").ToString))
                    'Me.spdRelease.SetText(Enum_Up.Item_Desc, Row, IIf(IsDBNull(sqlRdr("Description").ToString), "", sqlRdr("Description").ToString))

                    'Me.spdRelease.Row = Row
                    'Me.spdRelease.Col = Enum_Up.Sch_Qty
                    'Me.spdRelease.Text = IIf(IsDBNull(SCHQTY), 0, SCHQTY)

                    'Me.spdRelease.SetText(Enum_Up.CONSIGNEE_CODE, Row, IIf(Len(Me.txtCustomerCode.Text) = 0, "", Me.txtCustomerCode.Text))
                    'spdRelease.set_ColWidth(Enum_Up.SftyDays, 0)
                    'spdRelease.set_ColWidth(Enum_Up.SftyStkPerDay, 0)
                    'spdRelease.set_ColWidth(Enum_Up.wh_code, 0)
                    'spdRelease.set_ColWidth(Enum_Up.CONSIGNEE_CODE, 0)

                    'If Upload_FileType = "EDIFACT" Then
                    '    Me.spdRelease.SetText(Enum_Up.RAN_No, Row, IIf(IsDBNull(sqlRdr("RAN_NO")), "", sqlRdr("RAN_NO")))
                    '    spdRelease.set_ColWidth(Enum_Up.RAN_No, 0)
                    'End If

SKIP:
                    sql = String.Empty
                    'sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                    '    " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                    '    " Shipment_Qty,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                    '    " Updt_Dt,Updt_Uid,CallOffNoResetRemarks,dailypullflag,CONSIGNEE_CODE,BAG_QTY,transitDays, bufferDays , UNIT_CODE )" & _
                    '    " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdr("PARTY_ID1").ToString & "', " & _
                    '    " '" & Trim(sqlRdr("Cust_DrgNo").ToString) & "',Convert(DateTime,'" & sqlRdr("DELIVERY_DT").ToString & "',103)," & _
                    '    " '" & Val(sqlRdr("shipqty").ToString) & "',Convert(DateTime,'" & sqlRdr("DELIVERY_DT").ToString & "',103)," & _
                    '    " '" & SCHQTY & "',0,0," & _
                    '    " '" & Val(sqlRdr("shipqty").ToString) & "',getDate(),'" & mP_User & "',getDate()," & _
                    '    " '" & mP_User & "','" & Replace(Remarks, "'", "''") & "','0'," & _
                    '    " '" & txtCustomerCode.Text & "','" & mlngBAGQTY & "'," & Transit_Days & "," & Buffer_Days & " ,'" & gstrUNITID & "')"


                    sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                        " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                        " Shipment_Qty,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                        " Updt_Dt,Updt_Uid,CallOffNoResetRemarks,dailypullflag,CONSIGNEE_CODE,BAG_QTY,transitDays, bufferDays , UNIT_CODE )" & _
                        " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdr("PARTY_ID1").ToString & "', " & _
                        " '" & Trim(sqlRdr("Cust_DrgNo").ToString) & "',Convert(DateTime,'" & sqlRdr("DELIVERY_DT").ToString & "',103)," & _
                        " '" & Val(sqlRdr("shipqty").ToString) & "','" & getDateForDB(objDate) & "'," & _
                        " '" & SCHQTY & "',0,0," & _
                        " '" & Val(sqlRdr("shipqty").ToString) & "',getDate(),'" & mP_User & "',getDate()," & _
                        " '" & mP_User & "','" & Replace(Remarks, "'", "''") & "','0'," & _
                        " '" & txtCustomerCode.Text & "','" & mlngBAGQTY & "'," & Transit_Days & "," & Buffer_Days & " ,'" & gstrUNITID & "')"

                    SqlConnectionclass.ExecuteNonQuery(sql)

                    SCHQTY = 0

                End While

                sqlRdr.Close()
                sqlRdr = Nothing
                If Not objTrans Is Nothing Then
                    objTrans.Commit()
                    objTrans = Nothing
                End If
            Else
                MsgBox("Schedule Upload Required Flag is OFF for Customer " + txtCustomerCode.Text + vbCrLf + "Can't Upload File...")
                Return Nothing
                Exit Function
            End If

        Catch ex As Exception
            If Not objTrans Is Nothing Then
                objTrans.Rollback()
                objTrans = Nothing
            End If
            RaiseException(ex)
        End Try

    End Function

    Private Function FN_FORDSCHEDULE(ByVal TRANS_NUMBER As String) As Object

        Dim Transit_Days, Row, Buffer_Days As Integer

        Dim CustDrgNo As Object = Nothing, DELDT As Object = Nothing
        Dim SCHQTY As Double
        Dim SFTYDAYS_MNTD As Object = Nothing
        Dim sql As String = "", updSQL As String = "", strWHCode As String = ""
        Dim objDate As Object = Nothing

        Dim sqlRdr As SqlDataReader = Nothing

        Try
            If Upload_FileType = "FORD" Then

                sql = "SELECT DISTINCT DTL.DELIVERYDATE AS DELIVERY_DT ,DTL.CUST_DRGNO ,DTL.ITEM_CODE ,SUM(DTL.QTY) SHIPQTY " & _
                    " FROM SCHEDULE_UPLOAD_FORD_DTL DTL INNER JOIN SCHEDULE_UPLOAD_FORD_HDR HDR " & _
                    " ON HDR.Unit_Code = DTL.Unit_Code AND HDR.Doc_No = DTL.Doc_No " & _
                    " INNER JOIN SCHEDULEPARAMETER_MST SP ON SP.UNIT_CODE = HDR.Unit_Code AND SP.Customer_code = HDR.CustomerCode " & _
                    " INNER JOIN CUSTITEM_MST C ON C.UNIT_CODE = HDR.Unit_Code AND C.Account_Code = HDR.CustomerCode " & _
                    " AND C.Cust_Drgno = DTL.Cust_DrgNo INNER JOIN ITEM_MST I ON I.UNIT_CODE = C.UNIT_CODE " & _
                    " AND I.Item_Code = C.Item_code " & _
                    " WHERE HDR.CUSTOMERCODE  = '" & txtCustomerCode.Text & "' AND HDR.DOC_NO = " & txtDocNo.Text & "" & _
                    " AND HDR.UNIT_CODE = '" & gstrUnitId & "' " & _
                    " AND C.Active = 1 AND C.SCHUPLDREQD = 1 AND I.STATUS = 'A'" & _
                    " GROUP BY HDR.CUSTOMERCODE,HDR.UNIT_CODE ,HDR.DOC_NO ,DTL.CUST_DRGNO ,DTL.ITEM_CODE ,DTL.DELIVERYDATE ORDER BY DELIVERY_DT"

                sqlRdr = SqlConnectionclass.ExecuteReader(sql)

                SCHQTY = 0

                If sqlRdr.HasRows Then
                    sql = " Select TransitDaysBySea From ScheduleParameter_mst"
                    sql = sql & " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                    Transit_Days = SqlConnectionclass.ExecuteScalar(sql)

                    sql = " Select BufferDays From ScheduleParameter_mst"
                    sql = sql & " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                    Buffer_Days = SqlConnectionclass.ExecuteScalar(sql)

                    While sqlRdr.Read
                        mlngBAGQTY = 1

                        sql = "set dateformat dmy; select max(dt) as dt from Calendar_Mfg_mst" & _
                            " where work_flg=0 and dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), sqlRdr.Item("DELIVERY_DT")) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                        objDate = SqlConnectionclass.ExecuteScalar(sql)

                        SCHQTY = SCHQTY + Convert.ToDouble(sqlRdr.Item("shipqty"))

                        sql = String.Empty
                        sql = "select isnull(bag_qty,1) from item_mst where item_code = '" & sqlRdr.Item("Item_Code").ToString() & "' and active_flg = 1 and UNIT_CODE = '" & gstrUnitId & "'"
                        mlngBAGQTY = SqlConnectionclass.ExecuteScalar(sql)

                        If mlngBAGQTY = 0 Then mlngBAGQTY = 1

                        If SCHQTY > 0 Then
                            If Val(mlngBAGQTY) >= Val(SCHQTY) Then
                                SCHQTY = mlngBAGQTY
                            Else
                                If mlngBAGQTY > 0 Then
                                    If SCHQTY Mod mlngBAGQTY > 0 Then
                                        SCHQTY = (mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)) + mlngBAGQTY
                                    Else
                                        SCHQTY = mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)
                                    End If
                                End If
                            End If
                        End If

                        sql = String.Empty

                        sql = "set dateformat dmy; INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                           " Doc_No,Wh_code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                           " Shipment_Qty,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                           " Updt_Dt,Updt_Uid,CallOffNoResetRemarks,dailypullflag,CONSIGNEE_CODE,BAG_QTY,transitDays, bufferDays , UNIT_CODE ,internal_item_code)" & _
                           " VALUES('" & Trim(TRANS_NUMBER) & "','', " & _
                           " '" & Trim(sqlRdr.Item("Cust_DrgNo").ToString()) & "','" & Trim(sqlRdr.Item("DELIVERY_DT")) & "'," & _
                           " '" & Val(sqlRdr.Item("shipqty")) & "','" & getDateForDB(objDate) & "'," & _
                           " '" & SCHQTY & "',0,0," & _
                           " '" & Val(sqlRdr.Item("shipqty")) & "',getDate(),'" & mP_User & "',getDate()," & _
                           " '" & mP_User & "','" & Replace(Remarks, "'", "''") & "','0'," & _
                           " '" & txtCustomerCode.Text & "','" & mlngBAGQTY & "'," & Transit_Days & "," & Buffer_Days & " ,'" & gstrUnitId & "','" & Trim(sqlRdr.Item("ITEM_CODE").ToString()) & "')"

                        SqlConnectionclass.ExecuteNonQuery(sql)

                        SCHQTY = 0
                    End While
                End If

                sqlRdr.Close()
                sqlRdr = Nothing
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Function

    Private Function MoveFile() As Object

        On Error GoTo ERR_Renamed

        Dim FSO As New Scripting.FileSystemObject
        Dim file As String = Nothing
        Dim sql As String = Nothing
        Dim upldFiles As Scripting.File
        Dim folderName As String = ""
        Dim filearray(0) As Object
        Dim filedate(0) As Object
        Dim latestFile As String = ""
        Dim rsloc As New ClsResultSetDB
        Dim bkpLocation As String = ""
        Dim YesNo As String = ""
        Dim status As String = ""

        If mblnfilemove = True Then
            Return Nothing
            Exit Function
        End If

        Dim subFileName As String = ""
        mblnfilemove = False
        sql = "select TOP 1 BackUpLocation from scheduleparameter_mst" & " WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' and unit_code= '" & gstrUNITID & "' ORDER BY UPDDT DESC "
        rsloc.GetResult(sql)
        bkpLocation = rsloc.GetValue("BackUpLocation")
        rsloc.ResultSetClose()
        folderName = Mid(txtFileName.Text, 1, Len(txtFileName.Text) - InStr(1, StrReverse(txtFileName.Text), "\"))

        Obj_FSO = Nothing
        Obj_FSO = New Scripting.FileSystemObject
        Obj_FSO.GetFolder(folderName).Attributes = Scripting.FileAttribute.Normal

        If Obj_FSO.GetFolder(folderName).Files.Count > 0 Then
            For Each upldFiles In Obj_FSO.GetFolder(folderName).Files
                ReDim Preserve filearray(UBound(filearray) + 1)
                filearray(UBound(filearray)) = Mid(upldFiles.Path, Len(Obj_FSO.GetFolder(folderName).Path) + 2, Len(upldFiles.Path))

                If OptReleaseFile.Checked = True Then
                    If txtDocNo.Text = "" Then
                        Return Nothing
                        Exit Function
                    End If
                End If

                If Not FSO.FolderExists(bkpLocation) Then
                    'FSO.CreateFolder(bkpLocation).Attributes = System
                End If

                If bkpLocation = folderName Then
                    MsgBox("Source and Destination are Same", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Return Nothing
                    Exit Function
                End If

                If OptReleaseFile.Checked = True Then
                    file = bkpLocation & "\" & filearray(UBound(filearray))
                Else
                    file = bkpLocation & "\" & filearray(UBound(filearray))
                End If

                If FSO.FileExists(folderName & "\" & filearray(UBound(filearray))) = True Then
                    If FSO.FileExists(file) = True Then
                        FSO.DeleteFile(file, True)
                    End If

                    If UCase(folderName & "\" & filearray(UBound(filearray))) = UCase(txtFileName.Text) Then
                        status = "U"
                    Else
                        status = "M"
                    End If

                    subFileName = filearray(UBound(filearray))

                    'KillExcelProcess(Obj_EX)
                    FSO.MoveFile(folderName & "\" & subFileName, bkpLocation & "\")
                    mP_Connection.Execute("INSERT INTO BACKUPFILEHISTORY(" & " FILENAME, FILEDATE, STATUS,UNIT_CODE)" & " VALUES('" & filearray(UBound(filearray)) & "'," & " getDate(),'" & status & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Else
                    MsgBox("Source path does not exist", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            Next upldFiles

            MsgBox("Transaction completed Successfully.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
            mblnfilemove = True
            'Added for Issue ID eMpro-20080716-20353 Starts
            mblnDailymktUpdated = True
            'Added for Issue ID eMpro-20080716-20353 Ends
        End If
        Return Nothing
        Exit Function
ERR_Renamed:
        If Err.Number = 70 Then
            MsgBox("Backup Location is ReadOnly.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Function
        End If
        If Err.Number = 76 Then
            MsgBox("BackUp Location Not Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Function
        End If
        If Err.Number = 5 Then
            MsgBox("File Already Open, Cann't Move.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Function
        End If

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        Exit Function

    End Function



    'Private Function MoveFile() As Object

    '    Try

    '        Dim FSO As New Scripting.FileSystemObject
    '        Dim file As String = Nothing
    '        Dim sql As String = Nothing
    '        Dim upldFiles As Scripting.File
    '        Dim folderName As String = ""
    '        Dim filearray(0) As Object
    '        Dim filedate(0) As Object
    '        Dim latestFile As String = ""

    '        Dim sqlRdr As SqlDataReader 'rsloc
    '        Dim bkpLocation As String = ""
    '        Dim YesNo As String = ""
    '        Dim status As String = ""

    '        If mblnfilemove = True Then
    '            Return Nothing
    '            Exit Function
    '        End If

    '        Dim subFileName As String = ""
    '        mblnfilemove = False
    '        sql = "select TOP 1 BackUpLocation from scheduleparameter_mst" & " WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' and unit_code= '" & gstrUnitId & "' ORDER BY UPDDT DESC "
    '        sqlRdr = SqlConnectionclass.ExecuteReader(sql)
    '        If sqlRdr.HasRows Then
    '            While sqlRdr.Read
    '                bkpLocation = sqlRdr("BackUpLocation").ToString
    '            End While
    '        End If



    '        folderName = Mid(txtFileName.Text, 1, Len(txtFileName.Text) - InStr(1, StrReverse(txtFileName.Text), "\"))

    '        Obj_FSO = Nothing
    '        Obj_FSO = New Scripting.FileSystemObject
    '        Obj_FSO.GetFolder(folderName).Attributes = Scripting.FileAttribute.Normal

    '        If Obj_FSO.GetFolder(folderName).Files.Count > 0 Then
    '            For Each upldFiles In Obj_FSO.GetFolder(folderName).Files
    '                ReDim Preserve filearray(UBound(filearray) + 1)
    '                filearray(UBound(filearray)) = Mid(upldFiles.Path, Len(Obj_FSO.GetFolder(folderName).Path) + 2, Len(upldFiles.Path))

    '                If OptReleaseFile.Checked = True Then
    '                    If txtDocNo.Text = "" Then
    '                        Return Nothing
    '                        Exit Function
    '                    End If
    '                End If

    '                If Not FSO.FolderExists(bkpLocation) Then
    '                    FSO.CreateFolder(bkpLocation).Attributes = Scripting.FileAttribute.Normal
    '                End If

    '                If bkpLocation = folderName Then
    '                    MsgBox("Source and Destination are Same", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
    '                    Return Nothing
    '                    Exit Function
    '                End If

    '                If OptReleaseFile.Checked = True Then
    '                    file = bkpLocation & "\" & filearray(UBound(filearray))
    '                Else
    '                    file = bkpLocation & "\" & filearray(UBound(filearray))
    '                End If

    '                If FSO.FileExists(folderName & "\" & filearray(UBound(filearray))) = True Then
    '                    If FSO.FileExists(file) = True Then
    '                        FSO.DeleteFile(file, True)
    '                    End If

    '                    If UCase(folderName & "\" & filearray(UBound(filearray))) = UCase(txtFileName.Text) Then
    '                        status = "U"
    '                    Else
    '                        status = "M"
    '                    End If

    '                    subFileName = filearray(UBound(filearray))

    '                    FSO.MoveFile(folderName & "\" & subFileName, bkpLocation & "\")
    '                    sql = ""
    '                    sql = "INSERT INTO BACKUPFILEHISTORY(" & " FILENAME, FILEDATE, STATUS,UNIT_CODE)" & " VALUES('" & filearray(UBound(filearray)) & "'," & " getDate(),'" & status & "','" & gstrUnitId & "')"
    '                    SqlConnectionclass.ExecuteNonQuery(sql)
    '                Else
    '                    MsgBox("Source path does not exist", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
    '                End If
    '            Next upldFiles

    '            MsgBox("Transaction completed Successfully.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
    '            mblnfilemove = True
    '            mblnDailymktUpdated = True
    '        End If
    '        Return Nothing
    '    Catch ex As Exception
    '        RaiseException(ex)
    '    End Try
    'End Function

    Private Sub FillReleaseDiffDetails()
        Try
            Dim strsql As String
            Dim intMaxLoop As Short
            Dim oCmd As ADODB.Command
            Dim rsObj As ADODB.Recordset
            Dim sqlRdr As SqlDataReader
            Dim intLoopCounter As Short
            With VaSDiff
                .MaxRows = 0
                .MaxCols = 5
                .Row = 0
                .Col = 1 : .Text = "Cust Item Code"
                .set_ColWidth(1, 15)
                .Row = 0
                .Col = 2 : .Text = "Item Code"
                .set_ColWidth(1, 15)
                .Row = 0
                .Col = 3 : .Text = "Release Date"
                .set_ColWidth(2, 10)
                .Row = 0
                .Col = 4 : .Text = "Prev Qty"
                .set_ColWidth(3, 10)
                .Row = 0
                .Col = 5 : .Text = "New Qty"
                .set_ColWidth(4, 10)
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With

            With VaSDiff
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 2
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 3
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 4
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 5
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
            End With

            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            oCmd = New ADODB.Command
            rsObj = New ADODB.Recordset
            With oCmd
                .let_ActiveConnection(mP_Connection)
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandTimeout = 0
                .CommandText = "Proposed_Schedule_Diff_Hilex"
                .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUnitId))
                .Parameters.Append(.CreateParameter("@Cust_code", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustomerCode.Text)))
                .Parameters.Append(.CreateParameter("@FILETYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Upload_FileType))
            End With

            oCmd = Nothing
            intLoopCounter = 0
            While Not rsObj.EOF
                intLoopCounter = intLoopCounter + 1
                VaSDiff.MaxRows = VaSDiff.MaxRows + 1
                Call VaSDiff.SetText(1, intLoopCounter, rsObj.Fields("Cust_drgno"))
                Call VaSDiff.SetText(2, intLoopCounter, rsObj.Fields("Item_code"))
                Call VaSDiff.SetText(3, intLoopCounter, rsObj.Fields("Schedule_date"))
                Call VaSDiff.SetText(4, intLoopCounter, Val(rsObj.Fields("Prev_Qty").Value.ToString))
                Call VaSDiff.SetText(5, intLoopCounter, Val(rsObj.Fields("New_qty").Value.ToString))
                rsObj.MoveNext()
            End While
            VaSDiff.MaxRows = intLoopCounter
            rsObj.Close()
            rsObj = Nothing

            If intLoopCounter > 0 Then
                FramDiff.Visible = True
            Else
                FramDiff.Visible = False
            End If
            With VaSDiff
                .BlockMode = True
                .Row = 1
                .Row2 = intLoopCounter
                .Col = 0
                .Col2 = .MaxCols : .ForeColor = System.Drawing.Color.Blue
                .Lock = True
                .BlockMode = False
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FN_FILESELECTION()
        Try
            Dim folderName As String = Nothing
            Dim upldFiles As Scripting.File
            Dim FolderFiles As Scripting.File
            Dim filearray(0) As Object
            Dim upldFileName(0) As Object
            Dim filedate(0) As Object
            Dim latestFile As String = ""
            Dim YesNo As String = Nothing
            Dim Temp As String = Nothing
            Dim sqlRdr As SqlDataReader  'Rs
            Dim strsql As String = String.Empty
            Dim strConsList As String = String.Empty
            Dim strMaxWhDate As String = String.Empty

            spdRelease.MaxRows = 0

            If mShipmentFlag = True Then
                strsql = ""
                strsql = "select max(trans_dt) AS TRANS_DT from warehouse_stock_dtl where customer_code = '" & txtCustomerCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "'"

                strMaxWhDate = SqlConnectionclass.ExecuteScalar(strsql)
                strMaxWhDate = strMaxWhDate.Format("dd MMM yyyy")

                strsql = ""
                strsql = "select distinct consignee_code from warehouse_stock_dtl" & _
                    " where consignee_code not in (" & _
                    " select distinct(consignee_code) from warehouse_stock_dtl" & _
                    " where customer_code = '" & txtCustomerCode.Text & "'and UNIT_CODE= '" & gstrUnitId & "'  and trans_dt = (" & _
                    " select max(trans_dt) from warehouse_stock_dtl where customer_code = '" & txtCustomerCode.Text & "'and UNIT_CODE= '" & gstrUnitId & "')) AND CUSTOMER_CODE = '" & txtCustomerCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "'"

                sqlRdr = SqlConnectionclass.ExecuteReader(strsql)
                If sqlRdr.HasRows Then
                    While sqlRdr.Read()
                        strConsList = strConsList + sqlRdr("consignee_code").ToString + vbCrLf
                        sqlRdr.NextResult()
                    End While
                End If
                sqlRdr.Close()
                sqlRdr = Nothing

                If Trim(strConsList).Length > 0 Then
                    MsgBox("Warehouse Stock Not Uploaded For Following Consignees Of " & vbCrLf & UCase(txtCustomerCode.Text) & " On " & strMaxWhDate & vbCrLf & vbCrLf & strConsList & vbCrLf & "Upload Warehouse Stock For These Consignees.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If

            End If


            If OptReleaseFile.Checked Then
                strsql = "select top 1 ReleaseFile_Location from scheduleparameter_mst" & _
                    " where customer_code = '" & txtCustomerCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "' order by entdt"
                txtFileName.Text = SqlConnectionclass.ExecuteScalar(strsql)
            End If

            If Len(LTrim(RTrim(txtFileName.Text))) > 0 Then
                folderName = txtFileName.Text

                Temp = Mid(StrReverse(txtFileName.Text), 1, InStr(1, StrReverse(txtFileName.Text), "\") - 1)
                Obj_FSO = New Scripting.FileSystemObject
                If InStr(1, Temp, ".") = 0 Then
                    If Obj_FSO.FolderExists(folderName) = False Then
                        MsgBox("Folder Does Not Exist")
                        Exit Sub
                    End If
                Else
                    If Obj_FSO.FileExists(txtFileName.Text) = False Then
                        If spdRelease.MaxRows <= 1 Then
                            MsgBox("No Call-Offs present in the Release Folder.")
                        End If
                        Exit Sub
                    End If
                    folderName = VB.Left(folderName, Len(folderName) - Len(Temp) - 1)
                End If


                If Obj_FSO.GetFolder(folderName).Files.Count > 0 Then

                    If Obj_FSO.GetFolder(folderName).Files.Count > 1 Then
                        MsgBox("You Have More Than One File To Upload...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    End If
                    If Trim(latestFile) <> "" Then
                        txtFileName.Text = Obj_FSO.GetFolder(folderName).Path & "\" & latestFile ''& ".csv"
                    End If


                    ' Mayur
                    For Each upldFiles In Obj_FSO.GetFolder(folderName).Files
                        If Obj_FSO.GetFolder(folderName).Files.Count > 1 Then

                            If Path.GetExtension(upldFiles.Path) = ".862" Or Path.GetExtension(upldFiles.Path) = ".830" Then
                                latestFile = upldFiles.Path
                                txtFileName.Text = latestFile
                                latestFile = StrReverse(Mid(StrReverse(latestFile), 1, InStr(1, StrReverse(latestFile), "\") - 1))

                                FN_Release_File_Upload()
                            Else
                                ReDim Preserve filearray(UBound(filearray) + 1)
                                ReDim Preserve filedate(UBound(filedate) + 1)

                                filearray(UBound(filearray)) = Mid(upldFiles.Path, Len(Obj_FSO.GetFolder(folderName).Path) + 2, Len(upldFiles.Path))

                                filearray(UBound(filearray)) = VB.Left(filearray(UBound(filearray)), Len(filearray(UBound(filearray))))
                                If UBound(filedate) > 1 Then

                                    filedate(UBound(filedate) - 1) = Mid(filearray(UBound(filearray) - 1), Len(filearray(UBound(filearray) - 1)) - 22, 4) & Mid(filearray(UBound(filearray) - 1), Len(filearray(UBound(filearray) - 1)) - 17, 2) & Mid(filearray(UBound(filearray) - 1), Len(filearray(UBound(filearray) - 1)) - 14, 2) & Mid(filearray(UBound(filearray) - 1), Len(filearray(UBound(filearray) - 1)) - 11, 2) & Mid(filearray(UBound(filearray) - 1), Len(filearray(UBound(filearray) - 1)) - 8, 2) & Mid(filearray(UBound(filearray) - 1), Len(filearray(UBound(filearray) - 1)) - 5, 2)

                                    filedate(UBound(filedate)) = Mid(filearray(UBound(filearray)), Len(filearray(UBound(filearray))) - 22, 4) & Mid(filearray(UBound(filearray)), Len(filearray(UBound(filearray))) - 17, 2) & Mid(filearray(UBound(filearray)), Len(filearray(UBound(filearray))) - 14, 2) & Mid(filearray(UBound(filearray)), Len(filearray(UBound(filearray))) - 11, 2) & Mid(filearray(UBound(filearray)), Len(filearray(UBound(filearray))) - 8, 2) & Mid(filearray(UBound(filearray)), Len(filearray(UBound(filearray))) - 5, 2)

                                    If filedate(UBound(filedate)) < filedate(UBound(filedate) - 1) Then
                                        filearray(UBound(filearray)) = filearray(UBound(filearray) - 1)
                                    End If

                                    If filedate(UBound(filedate)) = filedate(UBound(filedate) - 1) Then
                                        MsgBox("There Are More Than One Release Files of Same Date to Upload", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                        Exit Sub
                                    End If
                                End If
                                latestFile = upldFiles.Path
                                latestFile = StrReverse(Mid(StrReverse(latestFile), 1, InStr(1, StrReverse(latestFile), "\") - 1))
                                txtFileName.Text = Obj_FSO.GetFolder(folderName).Path & "\" & latestFile ''& ".csv"
                                Call FN_Release_File_Upload() '' here mayur
                            End If

                        Else
                            latestFile = upldFiles.Path
                            latestFile = StrReverse(Mid(StrReverse(latestFile), 1, InStr(1, StrReverse(latestFile), "\") - 1))
                            txtFileName.Text = Obj_FSO.GetFolder(folderName).Path & "\" & latestFile ''& ".csv"
                            Call FN_Release_File_Upload() '' here mayur			
                        End If
                    Next

                    ' Mayur

                End If
            End If

            If Upload_FileType <> "EDIFACT" And Upload_FileType <> "COVISINT" And Upload_FileType <> "FORD" Then
                MsgBox("Wrong File Type-It's not EDIFACT/COVISINT/FORD format.", MsgBoxStyle.OkOnly, ResolveResString(100))
                Exit Sub
            End If
            If Upload_FileType = "FORD" Then
                Call MoveFile()
            End If

        Catch e As System.IO.FileNotFoundException
            MsgBox("Invalid File Name.", MsgBoxStyle.OkOnly, ResolveResString(100))
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Sub

    Private Sub FillMktScheduleDiffDetails(ByRef arrSchdDiff() As String)
        Try


            Dim intLoopCounter As Short

            Dim arrSchdRevision() As String


            With VaSDiff

                .MaxRows = 0
                .MaxCols = 5
                .Row = 0
                .Col = 1 : .Text = "Cust Item Code"
                .set_ColWidth(1, 15)
                .Row = 0
                .Col = 2 : .Text = "Item Code"
                .set_ColWidth(1, 15)
                .Row = 0

                .Col = 3 : .Text = "Schedule Date"
                .set_ColWidth(2, 10)
                .Row = 0
                .Col = 4 : .Text = "Prev Qty"
                .set_ColWidth(3, 10)
                .Row = 0
                .Col = 5 : .Text = "New Qty"
                .set_ColWidth(4, 10)
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With
            With VaSDiff
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 2
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 3
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 4
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
                .Row = .MaxRows
                .Col = 5
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
            End With

            intLoopCounter = 0
            For intLoopCounter = 0 To UBound(arrSchdDiff) - 1
                arrSchdRevision = Split(arrSchdDiff(intLoopCounter), ",")
                VaSDiff.MaxRows = VaSDiff.MaxRows + 1
                Call VaSDiff.SetText(1, intLoopCounter + 1, arrSchdRevision(0))
                Call VaSDiff.SetText(2, intLoopCounter + 1, arrSchdRevision(1))
                Call VaSDiff.SetText(3, intLoopCounter + 1, arrSchdRevision(3))
                Call VaSDiff.SetText(4, intLoopCounter + 1, arrSchdRevision(4).ToString)
                Call VaSDiff.SetText(5, intLoopCounter + 1, arrSchdRevision(5).ToString)
            Next

            VaSDiff.MaxRows = intLoopCounter

            If UBound(arrSchdDiff) > 0 Then
                FramDiff.Visible = True
            Else
                FramDiff.Visible = False
            End If
            FramDiff.Text = "&PLAUSIBILITY CHECK-Difference in Previous and Current Schedules for Planning"
            With VaSDiff
                .BlockMode = True
                .Row = 1
                .Row2 = UBound(arrSchdDiff)
                .Col = 0
                .Col2 = .MaxCols : .ForeColor = System.Drawing.Color.Blue
                .Lock = True
                .BlockMode = False
            End With

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Function FN_CheckTransmissionNo() As Boolean
        Try
            Dim FileType As String
            Dim strArr(10) As Object
            Dim YesNo As String
            Dim i As Short
            Dim j As Short
            Dim k As Short
            Dim Flag As Short


            ' mayur 
            Dim sqlRdr As SqlDataReader ' rsLASTUPLDCALLOFF
            Dim sqlRdr2 As SqlDataReader ' rsMissing
            Dim strSql As String = String.Empty

            Obj_EX = New Excel.Application
            Obj_EX.Workbooks.Open(txtFileName.Text)

            Remarks = ""

            If Replace(IIf(Obj_EX.Range("$A$1").Value Is Nothing, "", Obj_EX.Range("$A$1").Value.ToString), "'", "") = "DELFOR" Then
                Upload_FileType = "EDIFACT"


                strSql = "SELECT TOP 1 MSG_NUMBER,ENT_DT" & " From SCHEDULE_UPLOAD_DARWINEDIFACT_HDR" & " WHERE CUST_CODE = '" & txtCustomerCode.Text & "'" & " AND SENDERID = '" & Replace(IIf(Obj_EX.Range("$B$1").Value Is Nothing, "", Obj_EX.Range("$B$1").Value.ToString), "'", "") & "' " & " and unit_code='" & gstrUnitId & "' ORDER BY ENT_DT DESC"
                sqlRdr = SqlConnectionclass.ExecuteReader(strSql)

                If sqlRdr.HasRows = 0 Then

                    FN_CheckTransmissionNo = True
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Function

                End If

                If sqlRdr.HasRows Then
                    While sqlRdr.Read

                        If Val(Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "")) < Val(sqlRdr("MSG_NUMBER")) Then
                            Remarks = InputBox("Last Uploaded CalloffNo is: " & sqlRdr("MSG_NUMBER") & vbCrLf & "And Current CalloffNo is: " & Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "") & vbCrLf & "Do You Want to Reset the CalloffNo Series?" & vbCrLf & "If YES Enter Remarks - REMARKS ARE MANDATORY TO UPLOAD FILE.", "Reset CallOffNo Series")

                            If Len(LTrim(RTrim(Remarks))) <= 0 Then
                                FN_CheckTransmissionNo = False
                                If Not Obj_EX Is Nothing Then
                                    KillExcelProcess(Obj_EX)
                                    Obj_EX = Nothing
                                End If

                                Exit Function
                            End If

                            If frmMKTTRN0054.upld = True Then
                                FN_CheckTransmissionNo = True
                                If Not Obj_EX Is Nothing Then
                                    KillExcelProcess(Obj_EX)
                                    Obj_EX = Nothing
                                End If

                                frmMKTTRN0054.Dispose()
                                Exit Function
                            End If

                            FN_CheckTransmissionNo = True

                            strSql = String.Empty
                            strSql = "SELECT TOP 1 CALLOFFNO FROM AUTHCALLOFFS_HDR" & " WHERE LastCallOff = " & sqlRdr("MSG_NUMBER") & " " & " AND CUSTOMER_CODE = '" & txtCustomerCode.Text & "'" & " AND SENDERID = '" & Replace(IIf(Obj_EX.Range("$B$1").Value Is Nothing, "", Obj_EX.Range("$B$1").Value.ToString), "'", "") & "'" & " AND isnull(STATUS,'')='' and unit_code='" & gstrUnitId & "'"
                            sqlRdr2 = SqlConnectionclass.ExecuteScalar(strSql)

                            If sqlRdr2.HasRows Then
                                While sqlRdr2.Read
                                    GoTo MISSINGEDI
                                End While
                            End If
                        Else
                            If frmMKTTRN0054.upld = True Then
                                FN_CheckTransmissionNo = True
                                If Not Obj_EX Is Nothing Then
                                    KillExcelProcess(Obj_EX)
                                    Obj_EX = Nothing
                                End If

                                Exit Function
                            End If
                        End If


                        If Val(Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "")) = Val(sqlRdr("MSG_NUMBER")) Then
                            YesNo = CStr(MsgBox("This is the Last Uploaded CalloffNo, Do You Want to Upload It Again?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, ResolveResString(100)))
                            If YesNo = CStr(MsgBoxResult.Yes) Then
                                FN_CheckTransmissionNo = True
                                If Not Obj_EX Is Nothing Then
                                    KillExcelProcess(Obj_EX)
                                    Obj_EX = Nothing
                                End If

                                Exit Function
                            Else
                                FN_CheckTransmissionNo = False
                                If Not Obj_EX Is Nothing Then
                                    KillExcelProcess(Obj_EX)
                                    Obj_EX = Nothing
                                End If

                                Exit Function
                            End If
                        End If

                        If Val(Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "")) > Val(sqlRdr("MSG_NUMBER")) + 1 Then
MISSINGEDI:
                            YesNo = CStr(MsgBox("Missing Calloff: " & vbCrLf & "Last Uploaded CalloffNo: " & sqlRdr("MSG_NUMBER") & vbCrLf & "Current CalloffNo: " & Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "") & vbCrLf & "Can't Upload File...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100)))

                            strSql = String.Empty
                            strSql = "INSERT INTO AUTHCALLOFFS_HDR (CallOffNo," & " CallOffDate,SenderID,LastCallOff,Customer_Code,Customer_name," & " " & " FileName,IPADDRESS,Ent_Dt,Ent_UID,Upd_Dt,Upd_UID,unit_code)" & " VALUES " & " ('" & Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "") & "','" & FN_Date_Conversion(Trim(IIf(Obj_EX.Range("$D$1").Value Is Nothing, "", Obj_EX.Range("$D$1").Value.ToString))) & "', " & " '" & Replace(IIf(Obj_EX.Range("$B$1").Value Is Nothing, "", Obj_EX.Range("$B$1").Value.ToString), "'", "") & "' ,'" & sqlRdr("MSG_NUMBER").ToString & "','" & txtCustomerCode.Text & "'," & " '" & LTrim(RTrim(LblCustomerName.Text)) & "','" & LTrim(RTrim(txtFileName.Text)) & "','" & gstrIpaddressWinSck & "'," & " getDate(),'" & mP_User & "',getDate(),'" & mP_User & "','" & gstrUnitId & "')"
                            SqlConnectionclass.ExecuteNonQuery(strSql)

                            i = 1 : k = 0 : Flag = 0

                            Dim cell_data2 As String
range1:

                            If Obj_EX.Range("$A$" & i).Value Is Nothing Then
                                cell_data2 = ""
                            Else
                                cell_data2 = Obj_EX.Range("$A$" & i).Value.ToString()
                            End If
                            While cell_data2 <> ""
                                For j = 0 To k
                                    If strArr(j) = Replace(IIf(Obj_EX.Range("$V$" & i).Value Is Nothing, "", Obj_EX.Range("$V$" & i).Value.ToString), "'", "") Then
                                        Flag = 1
                                    End If
                                Next
                                If Flag = 0 Then
                                    strArr(k) = Replace(IIf(Obj_EX.Range("$V$" & i).Value Is Nothing, "", Obj_EX.Range("$V$" & i).Value.ToString), "'", "")
                                    k = k + 1
                                End If
                                i = i + 1
                                Flag = 0
                                GoTo range1
                            End While

                            For j = 0 To k - 1
                                strSql = String.Empty
                                strSql = "INSERT INTO AUTHCALLOFFS_DTL( CallOffNo," & " Factory_Code,WH_Code,SenderID,Consignee_Code,Consignee_Name,Ent_Dt,Ent_UID,Upd_Dt,Upd_UID,unit_code)" & " VALUES ('" & Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "") & "',''," & " '" & Replace(IIf(Obj_EX.Range("$V$" & j + 1).ToString Is Nothing, "", Obj_EX.Range("$V$" & j + 1).ToString), "'", "") & "','" & Replace(IIf(Obj_EX.Range("$B$1").Value Is Nothing, "", Obj_EX.Range("$B$1").Value.ToString), "'", "") & "'," & " '','',GETDATE()," & " '" & mP_User & "',GETDATE(),'" & mP_User & "','" & gstrUnitId & "')"
                                SqlConnectionclass.ExecuteNonQuery(strSql)
                            Next

                            FN_CheckTransmissionNo = False
                            If Not Obj_EX Is Nothing Then
                                KillExcelProcess(Obj_EX)
                                Obj_EX = Nothing
                            End If

                            Exit Function

                        End If

                        If Val(Replace(IIf(Obj_EX.Range("$J$1").Value Is Nothing, "", Obj_EX.Range("$J$1").Value.ToString), "'", "")) = Val(sqlRdr("MSG_NUMBER").ToString) + 1 Then
                            FN_CheckTransmissionNo = True

                        End If
                    End While
                End If
            End If

            If UCase(LTrim(VB.Left(IIf(Obj_EX.Range("$A$1").Value Is Nothing, "", Obj_EX.Range("$A$1").Value.ToString), 8))) = "COVISINT" Then
                Upload_FileType = "COVISINT"
                FN_CheckTransmissionNo = True
            End If

            If frmMKTTRN0054.upld = True Then
                FN_CheckTransmissionNo = True
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If

                Exit Function
            End If


        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
    End Function

    Private Function FN_DAILYPULLQTY(ByVal LNGWHSTOCK As Object, ByVal lngISSUEDQTY As Object, ByVal lngRCVDQTY As Object, ByVal lngSCHQTY As Long, ByVal varWHCODE As Object, ByVal lngSAFETYDAYS As Long, ByVal StrItemCode As String, ByVal strCustDrgNo As String, ByVal Row As Integer, ByVal strFACTORY_CODE As String, ByVal mlngBAGQTY As Long) As String
        Dim SCHQTY As Long
        Dim lngdlypullqty As Long
        Dim strsql As String = String.Empty
        Dim DSWAREDTL As DataTable
        Try

            'FN_DAILYPULLQTY(varWhStock, varIssuedQty, varRcvdQty, SCHQTY, Adors!SafetyDays, varWHCODE, CStr(Adors!Item_Code), CInt(Row), CStr(Adors!FactoryCode))
            'Created By         : Shubhra Verma
            'Created On         : 04 Mar 2008 to 07 Mar 2008
            'Issue id           : eMpro-20080306-13517
            'Revision History   :1 - There should be a provision of using daily pull
            '                    qty from Warehouse_Stock_dtl as minimum
            '                    safety stock if daily pull qty check box is checked
            '                    in CDP form.



            strsql = "SELECT  RATE FROM WAREHOUSE_STOCK_DTL" & _
            " WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "'" & _
            " AND WAREHOUSE_CODE = '" & varWHCODE & "'" & _
            " and item_code = '" & strCustDrgNo & "'" & _
            " and  TRANS_DT  = (SELECT MAX(TRANS_DT)FROM WAREHOUSE_STOCK_DTL" & _
            " WHERE  WAREHOUSE_CODE = '" & varWHCODE & "' and UNIT_CODE= '" & gstrUnitId & "'" & _
            " and customer_code = '" & txtCustomerCode.Text & "')" & _
            " and consignee_code in (select customer_code from customer_mst" & _
            " where dock_code = '" & strFACTORY_CODE & "' and UNIT_CODE= '" & gstrUnitId & "'" & _
            " and cust_vendor_code = '" & varWHCODE & "')" & _
            " and revno =  (select max(revno)" & _
            " From warehouse_stock_dtl" & _
            " WHERE CUSTOMER_CODE='" & txtCustomerCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "'" & _
            " AND WAREHOUSE_CODE='" & varWHCODE & "'" & _
            " and consignee_code in (select customer_code from customer_mst" & _
            " where dock_code = '" & strFACTORY_CODE & "' and UNIT_CODE= '" & gstrUnitId & "'" & _
            " and cust_vendor_code = '" & varWHCODE & "')" & _
            " and TRANS_DT  = (SELECT MAX(TRANS_DT)FROM WAREHOUSE_STOCK_DTL" & _
            " WHERE  WAREHOUSE_CODE = '" & varWHCODE & "'" & _
            " and customer_code = '" & txtCustomerCode.Text & "' and UNIT_CODE= '" & gstrUnitId & "'))"

            DSWAREDTL = SqlConnectionclass.GetDataTable(strsql)
            If (DSWAREDTL.Rows.Count > 0) Then
                lngdlypullqty = IIf(DSWAREDTL.Rows(0).Item("RATE").ToString() = "", 0, DSWAREDTL.Rows(0).Item("RATE").ToString())
            Else
                lngdlypullqty = 0
            End If


            If chkDlyPullQty.Checked = True Then
                lngdlypullqty = lngdlypullqty * lngSAFETYDAYS
                If Val(LNGWHSTOCK) + Val(lngRCVDQTY) < 0 Then

                    If lngSCHQTY < 0 Then lngSCHQTY = -lngSCHQTY
                    SCHQTY = lngSCHQTY + lngdlypullqty
                Else
                    If Val(LNGWHSTOCK) + Val(lngRCVDQTY) > Val(Val(lngdlypullqty) + Val(lngISSUEDQTY)) Then

                        SCHQTY = 0
                    Else
                        SCHQTY = Val(lngdlypullqty) + Val(lngISSUEDQTY) - Val(LNGWHSTOCK) - Val(lngRCVDQTY)
                    End If
                End If

                If SCHQTY > 0 Then

                    If Val(mlngBAGQTY) >= Val(SCHQTY) Then
                        SCHQTY = mlngBAGQTY
                    Else
                        If mlngBAGQTY > 0 Then
                            If SCHQTY Mod mlngBAGQTY > 0 Then
                                SCHQTY = (mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)) + mlngBAGQTY
                            Else
                                SCHQTY = mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)
                            End If
                        End If
                    End If

                End If

                If (DSWAREDTL.Rows.Count > 0) Then
                    FN_DAILYPULLQTY = CStr(SCHQTY) + "*" + CStr(IIf(DSWAREDTL.Rows(0).Item("RATE").ToString() = "", 0, DSWAREDTL.Rows(0).Item("RATE").ToString()))
                Else
                    FN_DAILYPULLQTY = CStr(SCHQTY) + "*" + "0"
                End If

                DSWAREDTL.Dispose()
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_EX.Workbooks.Close()
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End Try
        'Return CStr(SCHQTY)
    End Function

    Private Sub txtNoOfMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtNoOfMonths.Validating
        Try
            If optAvgofNextMonths.Enabled = True And optAvgofNextMonths.Visible = True Then
                If Val(txtNoOfMonths.Text) <= 1 Then
                    MsgBox("No Of Months Must Be Greater Than 1", MsgBoxStyle.Information, ResolveResString(100))
                    txtNoOfMonths.Focus()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Function Upload_EDIFACT(ByVal Cell_Data As String, ByVal Row As Integer, ByVal ShipmentFlag As Boolean) As Boolean

        Dim Cell_Data1 As String = String.Empty
        Dim Data_Row() As String
        Dim Msg As String = String.Empty
        Dim col As Integer = 0
        Dim i As Integer = 0
        Dim strSQL As String = String.Empty
        Dim SqlRdr As SqlDataReader = Nothing
        Dim sch As Integer = 0
        Dim WA As String = String.Empty
        Dim HOLIDAY As String = String.Empty
        Dim CustDrgNo As Object
        Dim ItemCode As Object
        Dim trans_number As String = String.Empty
        Dim fin_year_notation As String = String.Empty
        Dim YesNo As String = String.Empty
        Dim objConn As SqlConnection = Nothing
        Dim objTrans As SqlTransaction = Nothing

        Try
            Msg = ""
            If Len(Cell_Data) < 10 Then
                col = 1 : i = 0
                Cell_Data = ""
                If Not range.Value Is Nothing Then
                    Cell_Data1 = (range.Value.ToString)
                Else
                    Cell_Data1 = ""
                End If
                While Cell_Data1 <> ""
                    Cell_Data = Cell_Data & Cell_Data1 & ","
                    col = col + 1
                    range = Obj_EX.Cells(Row, col)
                    If Not range.Value Is Nothing Then
                        Cell_Data1 = (range.Value.ToString)
                    Else
                        Cell_Data1 = ""
                    End If
                End While
            End If

            Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
            For i = 0 To UBound(Data_Row)
                Data_Row(i) = Trim(Replace(Data_Row(i), "'", ""))
            Next i

            strSQL = "select Fin_Year_Notation from Financial_Year_Tb where unit_code = '" & gstrUnitId & "' and GETDATE() between Fin_Start_date and Fin_end_date "
            If Not IsRecordExists(strSQL) Then
                MessageBox.Show("Financial Year Not Defined.", ResolveResString(100), MessageBoxButtons.OK)
                Return False
            End If

            strSQL = "select Current_No from documenttype_mst where unit_code = '" & gstrUnitId & "'" & _
               " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            If Not IsRecordExists(strSQL) Then
                MessageBox.Show("Series Not Defined.", ResolveResString(100), MessageBoxButtons.OK)
                Return False
            End If

            objConn = New SqlConnection
            objConn = SqlConnectionclass.GetConnection()
            objTrans = objConn.BeginTransaction
            Flag = 0

            strSQL = "select Fin_Year_Notation from Financial_Year_Tb where unit_code = '" & gstrUnitId & "' and GETDATE() between Fin_Start_date and Fin_end_date "
            fin_year_notation = SqlConnectionclass.ExecuteScalar(strSQL)

            strSQL = "select Current_No from documenttype_mst where unit_code = '" & gstrUnitId & "'" & _
                " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            trans_number = SqlConnectionclass.ExecuteScalar(strSQL)

            trans_number = Val(trans_number) + 1
            strSQL = "update documenttype_mst set Current_No = " & trans_number & "  where unit_code = '" & gstrUnitId & "'" & _
               " and GETDATE() between Fin_Start_date and Fin_end_date and Doc_Type = 302"
            SqlConnectionclass.ExecuteNonQuery(strSQL)

            trans_number = fin_year_notation + trans_number

            strSQL = " Insert Into Schedule_Upload_DarwinEDIFACT_Hdr(Doc_No,Doc_Type,Cust_Code," & _
                "Consignee_Code,Plant_c,Upload_File_Name,Upload_File_Type,SenderID," & _
                "RecipientID,Receipt_Dt,Receipt_Time,Control_no,Test_Indicator,Msg_code," & _
                "Msg_Name,Msg_Number,Msg_Version,Upload_DtQualifier,Upload_Dt," & _
                "Upload_DtFormatQualifier,Start_DtQualifier,Start_Dt,Start_DtFormatQualifier,End_DtQualifier,End_Dt,End_DtFormatQualifier," & _
                "Party_Qualifier1,Party_ID1,Agency_code1,Party_Qualifier2,PARTY_ID2," & _
                "Agency_code2,Process_Indicator,Party_Qualifier3,Party_ID3,Agency_code3," & _
                "Ent_Dt,Upd_Dt,Ent_UId,Upd_UId,RevNo,UNIT_CODE) " & _
                " Values (" & trans_number & ",302,'" & Trim(txtCustomerCode.Text) & "'," & " '','" & Trim(txtUnitCode.Text) & "'," & _
                "'" & Trim(txtFileName.Text) & "',''," & " '" & Trim(Data_Row(1)) & "','" & Trim(Data_Row(2)) & "'," & _
                " '" & getDateForDB(FN_Date_Conversion_edifact(Trim(Data_Row(3)))) & "'," & " " & Val(Data_Row(4)) & ",'" & Trim(Data_Row(5)) & "'," & _
                " " & Val(Data_Row(6)) & ",'" & Trim(Data_Row(7)) & "', '" & Trim(Data_Row(8)) & "','" & Val(Data_Row(9)) & "'," & _
                " '" & Val(Data_Row(10)) & "', '" & Val(Data_Row(11)) & "','" & getDateForDB(FN_Date_Conversion_edifact(Trim(Data_Row(12)))) & "'," & _
                " '" & Trim(Data_Row(13)) & "', '" & Trim(Data_Row(14)) & "','" & getDateForDB(FN_Date_Conversion_edifact(Trim(Data_Row(15)))) & "'," & _
                " '" & Trim(Data_Row(16)) & "', '" & Trim(Data_Row(17)) & "','" & getDateForDB(FN_Date_Conversion_edifact(Trim(Data_Row(18)))) & "'," & _
                " '" & Trim(Data_Row(19)) & "', '" & Trim(Data_Row(20)) & "','" & Trim(Data_Row(21)) & "'," & _
                " '" & Trim(Data_Row(22)) & "', '" & Trim(Data_Row(23)) & "','" & Trim(Data_Row(24)) & "'," & _
                " '" & Trim(Data_Row(25)) & "', '" & Trim(Data_Row(26)) & "','" & Trim(Data_Row(27)) & "'," & _
                " '" & Trim(Data_Row(28)) & "', '" & Trim(Data_Row(29)) & "',getDate(),getDate(), " & _
                " '" & mP_User & "','" & mP_User & "',0,'" & gstrUnitId & "') "

            SqlConnectionclass.ExecuteNonQuery(strSQL)

            While Len(Cell_Data) <> 0
                If Len(Cell_Data) < 10 Then
                    col = 1 : i = 0
                    Cell_Data = ""
                    range = Obj_EX.Cells(Row, col)
                    If Not range.Value Is Nothing Then
                        Cell_Data1 = Trim(range.Value.ToString)
                    Else
                        Cell_Data1 = ""
                    End If

                    While Cell_Data1 <> ""
                        Cell_Data = Cell_Data & Cell_Data1 & ","
                        col = col + 1
                        range = Obj_EX.Cells(Row, col)
                        If Not range.Value Is Nothing Then
                            Cell_Data1 = Trim(range.Value.ToString)
                        Else
                            Cell_Data1 = ""
                        End If
                    End While
                End If

                Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
                For i = 0 To UBound(Data_Row)
                    Data_Row(i) = Trim(Replace(Data_Row(i), "'", ""))
                Next i

                If Trim(Data_Row(30)) = "" Then
                    Flag = 1
                    MsgBox("Item Code is blank.File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Return False
                End If

                If FN_Date_Conversion_edifact(Trim(Data_Row(3))) = "" Then
                    Flag = 1
                    Return False
                ElseIf CDate(FN_Date_Conversion_edifact(Trim(Data_Row(3)))) > CDate(IIf(FN_Date_Conversion_edifact(Trim(Data_Row(46))) = "", "01/01/1900", FN_Date_Conversion_edifact(Trim(Data_Row(46))))) Then
                    If Trim(Data_Row(40)) <> "" And Trim(Data_Row(41)) <> "" Then
                        MsgBox("Schedule Date " + FN_Date_Conversion_edifact(Trim(Data_Row(46))) + " Should Be Greater Than" + vbCrLf + "Transmission Date " + FN_Date_Conversion_edifact(Trim(Data_Row(3))) + " of Release File" + vbCrLf + "File Can't Be Uploaded", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Flag = 1
                        Return False
                    End If
                End If

                Dim STRCONS As String

                strSQL = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST WHERE CUST_VENDOR_CODE = '" & Trim(Data_Row(21)) & "'" & _
                    " AND DOCK_CODE = '" & Trim(Data_Row(28)) & "' and UNIT_CODE='" & gstrUnitId & "'"

                If IsRecordExists(strSQL) Then
                    STRCONS = SqlConnectionclass.ExecuteScalar(strSQL)
                Else
                    STRCONS = ""
                End If

                strSQL = " Insert into Schedule_Upload_DarwinEDIFACT_dtl(Doc_No,Doc_Type, " & _
                " Cust_Code,Consignee_Code,Item_Code,Item_Type,ProductID_Qualifier,Item_Number," & _
                " Item_NumberType,Location_Qualifier,Location_ID,Ref_Qualifier,Ref_ID," & _
                " DelPlan_Status,Frequency,Dispatch_Pattern,Quantity_Qualifier,Quantity,UOM," & _
                " DelDT_Qualifier,Delivery_DT,DelDT_FormatQualifier,Ent_Dt,Upd_Dt,Ent_UId,Upd_UId,RevNo,RAN_No,UNIT_CODE) " & _
                " Values (" & trans_number & ",302,'" & Trim(txtCustomerCode.Text) & "'," & _
                " '" & Trim(STRCONS) & "','" & Trim(Data_Row(30)) & "','" & Trim(Data_Row(31)) & "', " & _
                " '" & Trim(Data_Row(32)) & "','" & Trim(Data_Row(33)) & "','" & Trim(Data_Row(34)) & "', " & _
                " '" & Trim(Data_Row(35)) & "','" & Trim(Data_Row(36)) & "','" & Trim(Data_Row(37)) & "', " & _
                " '" & Trim(Data_Row(38)) & "','" & Trim(Data_Row(39)) & "','" & Trim(Data_Row(40)) & "', " & _
                " '" & Trim(Data_Row(41)) & "','" & Trim(Data_Row(42)) & "'," & Val(Data_Row(43)) & ", " & _
                " '" & Trim(Data_Row(44)) & "','" & Trim(Data_Row(45)) & "','" & getDateForDB(FN_Date_Conversion_edifact(Trim(Data_Row(46)))) & "', " & _
                " '" & Trim(Data_Row(47)) & "',getDate(),getDate() ,'" & mP_User & "','" & mP_User & "'," & _
                " 0,'" & Trim(Data_Row(48)) & "','" & gstrUnitId & "')"
                SqlConnectionclass.ExecuteNonQuery(strSQL)

                Row = Row + 1

                range = Obj_EX.Cells(Row, 1)
                If Not range.Value Is Nothing Then
                    Cell_Data = (range.Value.ToString)
                Else
                    Cell_Data = ""
                End If
            End While

            Dim countREC As Short = 0

            strSQL = "select DISTINCT C.ACCOUNT_CODE, C.cust_drgno,COUNT(C.item_code) countitem from custitem_mst C with (nolock)" & _
                " where C.active = 1 AND SCHUPLDREQD = 1 and  unit_code='" & gstrUnitId & "' and "

            If ShipmentFlag = True Then
                strSQL = strSQL & " C.ACCOUNT_CODE IN (SELECT DISTINCT CONSIGNEE_CODE"
                strSQL = strSQL & " FROM SCHEDULE_UPLOAD_DARWINEDIFACT_DTL with (nolock) where UNIT_CODE ='" & gstrUnitId & "')" & _
                        " AND C.CUST_DRGNO IN (SELECT DISTINCT CUST_DRGNO"
            Else
                strSQL = strSQL & " C.ACCOUNT_CODE = '" & txtCustomerCode.Text & "'"
                strSQL = strSQL & " AND C.CUST_DRGNO IN (SELECT DISTINCT CUST_DRGNO"
            End If

            strSQL = strSQL & " FROM SCHEDULE_UPLOAD_DARWINEDIFACT_DTL with (nolock)"

            strSQL = strSQL & " WHERE DOC_NO = '" & trans_number & "'  and  UNIT_CODE ='" & gstrUnitId & "') " & _
                " group by C.Account_code, C.cust_drgno"

            SqlRdr = SqlConnectionclass.ExecuteReader(strSQL)

            Msg = ""

            If SqlRdr.HasRows Then
                While SqlRdr.Read
                    countREC = SqlRdr("COUNTITEM").ToString
                    If countREC > 1 Then
                        Msg = Msg & "  " + SqlRdr("cust_drgno").ToString
                        Flag = 1
                    End If
                End While
                If Msg <> "" Then
                    Flag = 1
                    MsgBox("Following Cust_DrgNo(s) Are Active For Multiple Items " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    Return False
                End If
            End If

            SqlRdr.Close()
            SqlRdr = Nothing

            strSQL = "select distinct d.ITEM_CODE from Schedule_Upload_DarwinEDIFACT_dtl d,Schedule_Upload_DarwinEDIFACT_hdr h " & _
                " Where d.Unit_Code=h.Unit_Code and h.revno = 0 and h.UNIT_CODE ='" & gstrUnitId & "' and h.cust_code = '" & Me.txtCustomerCode.Text & "'" & _
                " and d.doc_no = h.doc_no and h.doc_no=" & trans_number & " and ltrim(rtrim(d.ITEM_CODE)) not in (select cust_drgno " & _
                " from custitem_mst where active = 1 AND SCHUPLDREQD = 1 and  UNIT_CODE ='" & gstrUnitId & "'"

            If ShipmentFlag = True Then
                strSQL = strSQL + " and account_code in (SELECT DISTINCT CONSIGNEE_CODE FROM SCHEDULE_UPLOAD_DARWINEDIFACT_DTL WHERE DOC_NO = " & trans_number & " and  UNIT_CODE ='" & gstrUnitId & "'))"
            Else
                strSQL = strSQL & " and account_code = '" & Me.txtCustomerCode.Text & "')"
            End If

            SqlRdr = SqlConnectionclass.ExecuteReader(strSQL)

            Msg = ""
            If SqlRdr.HasRows Then
                While SqlRdr.Read
                    Msg = Msg & "'" + SqlRdr("ITEM_CODE").ToString + "'" + vbCrLf
                End While

                If Len(Trim(Msg)) > 0 Then
                    SqlRdr.Close()
                    SqlRdr = Nothing
                    MsgBox("Following Items Are Not Defined In The System" & vbCrLf & "for Customer " & txtCustomerCode.Text & " : " & vbCrLf & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                    Flag = 1
                    Return False
                End If
            End If

            HOLIDAY = Replace(HOLIDAY, " ", Chr(13))
            If HOLIDAY <> "" Then
                MsgBox("Following is/are not working day(s):" & vbCrLf & vbCrLf & "Consignee---Date" & vbCrLf & HOLIDAY & vbCrLf)
                Flag = 1
                Return False
            End If

            If Me.optWkgDays.Checked = True Then
                WA = "W"
            Else
                WA = "A"
            End If

            If Me.optCurMonthSch.Checked = True Then
                sch = 0
            ElseIf Me.optNextMonthSch.Checked = True Then
                sch = 1
            Else
                sch = Val(Me.txtNoOfMonths.Text)
            End If

            If ShipmentFlag = False And chkDlyPullQty.Checked = True Then
                MsgBox("Shipment For This Customer Is Not Through Warehouse, " & vbCrLf & "So, You Can't Use Daily Pull Qty For Safety Stock Calculations.", vbInformation, ResolveResString(100))
                Flag = 1
                Return False
            Else
                If Not objTrans Is Nothing Then
                    objTrans.Commit()
                    objTrans = Nothing
                End If
                Me.txtDocNo.Text = trans_number
            End If

            If ShipmentFlag = False Then
                Call FN_Display_WITHOUTWH(trans_number)
            End If

            strSQL = "Select Top 1 Doc_No From ScheduleProposalcalculations Where Doc_No= '" & Val(trans_number) & "'" & _
                " and shipment_qty > 0 AND UNIT_CODE = '" & gstrUnitId & "'"

            If Not IsRecordExists(strSQL) Then
                YesNo = CStr(MsgBox("No Schedule Proposed." + vbCrLf + "Do You Want To Move Files To BackUp Folder?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)))
                If YesNo = CStr(MsgBoxResult.Yes) Then Call MoveFile()
                If YesNo = CStr(MsgBoxResult.No) Then
                    Return True
                End If
            Else
                MsgBox(" Schedule has been Uploaded Succesfully !", MsgBoxStyle.Information, ResolveResString(100))
                Call Updt_DailyMkt(Upload_FileType)
                Call MoveFile()
            End If

            Return True

            SqlRdr.Close()
            SqlRdr = Nothing

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            If Not SqlRdr.IsClosed Then
                SqlRdr.Close() : SqlRdr = Nothing
            End If

            If Not objTrans Is Nothing Then
                If Flag = 1 Then
                    objTrans.Rollback()
                Else
                    objTrans.Commit()
                End If
                objTrans = Nothing
            End If
        End Try
    End Function

    Private Function Upload_ford(ByVal Cell_Data As String, ByVal Row As Integer, ByVal trans_number As Integer, ByVal file_extension As String) As Boolean

        Dim Cell_Data1 As String = String.Empty
        Dim Data_Row() As String
        Dim Msg As String = String.Empty
        Dim col As Integer = 0
        Dim i As Integer = 0
        Dim strSQL As String = String.Empty
        Dim SqlRdr As SqlDataReader = Nothing
        Dim sch As Integer = 0
        Dim WA As String = String.Empty
        Dim HOLIDAY As String = String.Empty
        Dim CustDrgNo As Object
        Dim ItemCode As Object
        Dim SftyStk As Object
        Dim SftyDays As Object
        Dim ShpgQty As Object
        Dim ShipDate As Object
        Dim wh_code As Object
        Dim CONSIGNEE_CODE As Object
        Dim YesNo As String = String.Empty

        Try
            Msg = ""
            If Len(Cell_Data) < 10 Then
                col = 1 : i = 0
                Cell_Data = ""
                If Not range.Value Is Nothing Then
                    Cell_Data1 = (range.Value.ToString)
                Else
                    Cell_Data1 = ""
                End If
                While Cell_Data1 <> ""
                    Cell_Data = Cell_Data & Cell_Data1 & ","
                    col = col + 1
                    range = Obj_EX.Cells(Row, col)
                    If Not range.Value Is Nothing Then
                        Cell_Data1 = (range.Value.ToString)
                    Else
                        Cell_Data1 = ""
                    End If
                End While
            End If

            Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
            For i = 0 To UBound(Data_Row)
                Data_Row(i) = Trim(Replace(Data_Row(i), "'", ""))
            Next i


            If file_extension = ".862" Then
                strSQL = "set dateformat dmy; Insert Into schedule_upload_ford_hdr(Doc_No,MsgNo,MsgDate,CustomerCode,UpldFileName,Ent_Dt,Ent_UserID,Upd_Dt,Upd_UserID,Unit_Code) " & _
                           " Values (" & trans_number & ",'" & Trim(Data_Row(0)) & "','" & FN_Date_Conversion_edifact(Trim(Data_Row(1))) & "','" & Trim(txtCustomerCode.Text) & "'," & _
                           "'" & Trim(txtFileName.Text) & "'," & _
                           " getDate(), " & _
                           " '" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUnitId & "') "
            End If

            If file_extension = ".830" Then
                strSQL = "set dateformat dmy; Insert Into schedule_upload_ford_hdr(Doc_No,Plant,CustomerCode,UpldFileName,Ent_Dt,Ent_UserID,Upd_Dt,Upd_UserID,Unit_Code) " & _
                          " Values (" & trans_number & ",'" & Trim(Data_Row(0)) & "','" & Trim(txtCustomerCode.Text) & "'," & _
                          "'" & Trim(txtFileName.Text) & "'," & _
                          " getDate(), " & _
                          " '" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUnitId & "') "

            End If

            SqlConnectionclass.ExecuteNonQuery(strSQL)

            While Len(Cell_Data) <> 0
                If Len(Cell_Data) < 10 Then
                    col = 1 : i = 0
                    Cell_Data = ""
                    range = Obj_EX.Cells(Row, col)
                    If Not range.Value Is Nothing Then
                        Cell_Data1 = Trim(range.Value.ToString)
                    Else
                        Cell_Data1 = ""
                    End If

                    While Cell_Data1 <> ""
                        Cell_Data = Cell_Data & Cell_Data1 & ","
                        col = col + 1
                        range = Obj_EX.Cells(Row, col)
                        If Not range.Value Is Nothing Then
                            Cell_Data1 = Trim(range.Value.ToString)
                        Else
                            Cell_Data1 = ""
                        End If
                    End While
                End If

                Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
                For i = 0 To UBound(Data_Row)
                    Data_Row(i) = Trim(Replace(Data_Row(i), "'", ""))
                Next i

                If Trim(Data_Row(2)) = "" Then
                    MsgBox("Item Code is blank.File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Return False
                End If

                If file_extension = ".862" Then
                    strSQL = "select item_code FROM CUSTITEM_MST " & _
                      " WHERE CUST_DRGNO = '" & Replace(Trim(Data_Row(2)), " ", "-") & "'" & " AND active = 1 AND SCHUPLDREQD = 1 AND" & _
                      " account_code = '" & Me.txtCustomerCode.Text & "' and UNIT_CODE='" & gstrUnitId & "'"

                    ItemCode = SqlConnectionclass.ExecuteScalar(strSQL)

                    strSQL = "set dateformat dmy; Insert into schedule_upload_ford_dtl(Doc_No,Cust_DrgNo,Item_Code,ShipFrom,ShipTo,Line_No,DeliveryDate,Qty,Ent_Dt,Upd_Dt,Ent_UserID,Upd_UserID,Unit_Code) " & _
                    " Values (" & trans_number & ",'" & Trim(Data_Row(2)) & "'," & _
                " '" & ItemCode & "','" & Trim(Data_Row(3)) & "', " & _
                " '" & Trim(Data_Row(4)) & "'," & Trim(Data_Row(5)) & ",'" & FN_Date_Conversion_edifact(Trim(Data_Row(6))) & "', " & _
                " " & Trim(Data_Row(7)) & ",getDate(),getDate() ,'" & mP_User & "','" & mP_User & "'," & _
                " '" & gstrUnitId & "')"

                End If

                If file_extension = ".830" Then

                    strSQL = "select item_code FROM CUSTITEM_MST " & _
                    " WHERE CUST_DRGNO = '" & Replace(Trim(Data_Row(1)), " ", "-") & "'" & " AND active = 1 AND SCHUPLDREQD = 1 AND" & _
                    " account_code = '" & Me.txtCustomerCode.Text & "' and UNIT_CODE='" & gstrUnitId & "'"

                    ItemCode = SqlConnectionclass.ExecuteScalar(strSQL)

                    strSQL = "set dateformat dmy; Insert into schedule_upload_ford_dtl(Doc_No,Cust_DrgNo,Item_Code,ShipFrom,Line_No,NoOfWeeks,DeliveryDate,Qty,Ent_Dt,Upd_Dt,Ent_UserID,Upd_UserID,Unit_Code) " & _
                    " Values (" & trans_number & ",'" & Trim(Data_Row(1)) & "'," & _
                    " '" & ItemCode & "','" & Trim(Data_Row(6)) & "', " & _
                    " '" & Trim(Data_Row(4)) & "'," & Trim(Data_Row(5)) & ",'" & Trim(Data_Row(3)) & "', " & _
                    " " & Trim(Data_Row(2)) & ",getDate(),getDate() ,'" & mP_User & "','" & mP_User & "'," & _
                    " '" & gstrUnitId & "')"
                End If

                SqlConnectionclass.ExecuteNonQuery(strSQL)

                Row = Row + 1

                range = Obj_EX.Cells(Row, 1)
                If Not range.Value Is Nothing Then
                    Cell_Data = (range.Value.ToString)
                Else
                    Cell_Data = ""
                End If
            End While

            Return True
        Catch ex As Exception
            RaiseException(ex)
            Return False
        End Try
    End Function
	Private Sub AutoMailerSend(ByVal DocNo As Integer)
        Dim FunReturn As Boolean = False
        Try
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_AUTOMAILER_SCHEDULE_UPLOADED"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 18).Value = DocNo.ToString()
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                End With
            End Using
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class