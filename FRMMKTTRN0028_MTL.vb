Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Friend Class FRMMKTTRN0028_MTL
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
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 26 APR 2010 to 04 MAY 2010
    'ISSUE ID           : eMpro-20100503-46268
    '                   : Split Customer Schedule into Daily Schedule.
    '                   : Upload WH Stock for ford in txt format.
    '********************************************************************************************************
    '============================================================================================
    'Revised By         :   Shalini Singh
    'Revised On         :   28 Sep 2011
    'Reason             :   Ip address and C drive hard codeed change
    'issue id           :   10140039
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 24 Oct 2011
    'ISSUE ID           : 10116406
    'Description        : 1 - the factory dispatch quantity is not showing in standard packing quantity
    '                     2 – Release File not moved to backup location.
    '============================================================================================
    'REVISED BY         : VIRENDRA GUPTA
    'REVISED ON         : 06 FEB 2012
    'Description        : MULTIUNIT CHANGES
    '============================================================================================
    'eMpro-20100503-46268
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 18 JUL 2012
    'ISSUE              : 10251161 - SOME ITEMS ARE NOT APPEARING IN CDP - MTL SHARJAH
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 03 AUG 2012
    'ISSUE              : 10275822 - Changes for MTL Sharjah
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 06 NOV 2012
    'ISSUE              : 10304880 
    'Description        : .xlsx format included for MTL Sharjah CDP.
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 10 JAN 2013
    'ISSUE              : 10329839 
    'Description        : Schedule Not Generating for some items through CDP
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 23 JAN 2013
    'ISSUE              : 10333808 
    'Description        : if schedule date is less than the minimum date in manufacturing calendar, system gives conversion error message
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 23 JUN 2015
    'ISSUE              : 10793737
    'Description        : CDP changes for MTL Sharjah
    '============================================================================================
    'REVISED BY         : SHUBHRA VERMA
    'REVISED ON         : 24 SEP 2015
    'ISSUE              : 10902693 
    'Description        : Calculating wrong WH Stock, when there are multiple rows of same date but different Quantity.
    '============================================================================================
    Dim mintFormIndex As Short
    Dim Obj_FSO As Scripting.FileSystemObject
    Dim Obj_EX As Excel.Application
    Dim Darwin_FileType As String
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
    Dim ShipmentFlag As Boolean
    Dim CONSIGNEE_CODE As String

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
    End Enum

    Private Enum BoschExcelColumn
        EDI_code = 1
        Call_Off_No
        Call_Off_Date
        Buyer_Plant_code
        PlantCode
        SupplierCode
        Buyer_Number
        Place_port
        Place_port_discharge
        Ref_Order_Number
        Pre_delivery_InstructionNo
        Cumulative_Quantity
        Cumulative_Quantity_Startdate
        Schedule_condition
        Dispatch_Qty
        Sche_Delivery_Date
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

#Region "Form Level Constant"
    Private Const HeaderRowIndex As Byte = 1
    Private Const DataRowIndex As Byte = 1
#End Region

#Region "Form Level Variable"
    Dim BoshExcelColumnName As String() = {"RcvrEDICode", "CallOffNo", "CallOffDate", "SupplyToBuyerPlantcode", "SupplyFromPlantCode", "SupplierCode", "BuyerPartNumber", "PortOfDischarge", "PortOfDischarge_AddIntDest", "ReferenceOrderNumber", "PrevDeliveryInstrNo", "CumQuantityReceived", "CumQtyStartDate", "SchCondition", "DispQty", "SchDeliveryDate", "Customer_Code", "Consignee_Code"}
#End Region

    Public Function FN_Date_Conversion(ByRef Cell_Dt As String) As Object
        On Error GoTo ERR_Renamed
        Dim T_Month, T_Date, T_Year As String
        Dim RSconsignee As New ClsResultSetDB
        Dim HOLIDAY As Short
        Dim dtDate As Date
        Dim strDate As String
        Dim strQry As String
        Dim rs As ClsResultSetDB

        HOLIDAY = 1
        If Cell_Dt = "555555" Or Cell_Dt = "444444" Then
            FN_Date_Conversion = ""
            Exit Function
        End If

        Cell_Dt = Replace(Cell_Dt, "'", "")
        If Len(Cell_Dt) >= 5 Then
            If Len(Cell_Dt) > 8 Then
                Cell_Dt = Mid(Cell_Dt, 1, 8)
            End If
            T_Date = Mid(Cell_Dt, Len(Cell_Dt) - 1, 2)
            T_Month = Mid(Cell_Dt, Len(Cell_Dt) - 3, 2)
            T_Year = Mid(Cell_Dt, 1, Len(Cell_Dt) - 4)
            If Len(T_Year) = 1 Then
                T_Year = "200" & T_Year
            ElseIf Len(T_Year) = 2 Then
                T_Year = "20" & T_Year
            End If
            If T_Date = "00" Then T_Date = "01"
            If Cell_Dt = "333333" Then
                T_Date = CDate(getDateForDB(GetServerDate()))
                strDate = T_Date
            ElseIf T_Month = "00" Then

                strQry = "Select * FROM DBO.FN_GETDATE_WEEKYEAR(" & T_Date & "," & T_Year & ")"
                rs = New ClsResultSetDB
                If rs.GetResult(strQry) = True Then
                    T_Date = CDate(getDateForDB(rs.GetValue("WEEKOFDAY")))
                    strDate = T_Date
                    rs.ResultSetClose()
                    rs = Nothing
                End If

            Else

                strDate = T_Date & "/" & T_Month & "/" & T_Year
            End If

            dtDate = VB6.Format(strDate, "dd/MMM/yyyy")
            If IsDate(getDateForDB(dtDate)) = True Then
                FN_Date_Conversion = getDateForDB(dtDate)
                If Mid(Cell_Dt, Len(Cell_Dt) - 1, 2) = "00" Then
                    mP_Connection.Execute("set dateformat 'dmy' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    RSconsignee.GetResult("select work_flg from calendar_mkt_cust" & " where dt = convert(varchar,'" & FN_Date_Conversion & "',106) " & " and Cust_Code = '" & consignee & "' AND UNIT_CODE ='" & gstrUNITID & "'")
                    While RSconsignee.GetValue("WORK_FLG") = True
                        T_Date = CStr(CDbl(T_Date) + 1)
                        strDate = CStr(CDbl(T_Date) + 1).PadLeft(2, "0") & "/" & T_Month & "/" & T_Year
                        FN_Date_Conversion = getDateForDB(ConvertToDate(strDate, "dd/MM/yyyy"))
                        RSconsignee.ResultSetClose()
                        RSconsignee = New ClsResultSetDB
                        RSconsignee.GetResult("select work_flg from calendar_mkt_cust" & " where dt = '" & FN_Date_Conversion & "'" & " and Cust_Code = '" & consignee & "' AND UNIT_CODE = '" & gstrUNITID & "'")
                    End While
                    RSconsignee.ResultSetClose()

                End If

            Else
                FN_Date_Conversion = ""
            End If
        Else
            dtDate = ConvertToDate("01/01/1900", "dd/MM/yyyy")
            FN_Date_Conversion = getDateForDB(dtDate)
        End If

            Exit Function
ERR_Renamed:
            If Err.Number = 13 Then
                FN_Date_Conversion = ""
                Exit Function
            End If
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Exit Function

    End Function

    Public Function FN_Date_Conversion_edifact(ByRef Cell_Dt As String) As Object

        On Error GoTo ERR_Renamed

        Dim T_Month, T_Date, T_Year As String
        Dim dtDate As Date
        Dim strDate As String

        Cell_Dt = Replace(Cell_Dt, "'", "")


        If Len(Cell_Dt) >= 5 Then
            If Len(Cell_Dt) > 8 Then
                Cell_Dt = Mid(Cell_Dt, 1, 8)
            End If
            T_Date = Mid(Cell_Dt, Len(Cell_Dt) - 1, 2)
            T_Month = Mid(Cell_Dt, Len(Cell_Dt) - 3, 2)
            T_Year = Mid(Cell_Dt, 1, Len(Cell_Dt) - 4)
            If Len(T_Year) = 1 Then
                T_Year = "200" & T_Year
            ElseIf Len(T_Year) = 2 Then
                T_Year = "20" & T_Year
            End If
            strDate = T_Date & "/" & T_Month & "/" & T_Year
            dtDate = ConvertToDate(strDate, "dd/MM/yyyy")
            If IsDate(getDateForDB(dtDate)) = True Then
                FN_Date_Conversion_edifact = getDateForDB(dtDate)
            Else
                FN_Date_Conversion_edifact = ""
            End If
        Else
            dtDate = ConvertToDate("01/01/1900", "dd/MM/yyyy")
            FN_Date_Conversion_edifact = getDateForDB(dtDate)
        End If

        Exit Function
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If
        Exit Function

    End Function

    Private Function FN_Display(ByVal TRANS_NUMBER As String, ByVal FileType As String) As String
        On Error GoTo ERR_Renamed
        'CHANGED BY SHUBHRA ON 10 APR 2008
        'ISSUE ID : eMpro-20080410-17008
        Dim Transit_Days, Row, Buffer_Days As Integer
        Dim SCHQTY As Long
        Dim sqlCMDDisplay As SqlCommand
        Dim sqlRdrDisplay As SqlDataReader
        Dim SFTYDAYS_MNTD As Object
        Dim SAFETYDAYS_BELOW As Long
        Dim TMPWHSTOCK As Long
        Dim rsDate As SqlCommand
        Dim sql As String = String.Empty, updSQL As String = String.Empty, strWHCode As String = String.Empty
        Dim RSWHSTOCK As New ClsResultSetDB
        Dim rsWHDt As SqlCommand
        Dim rdrWHDt As SqlDataReader
        Dim varWhStock As Object = Nothing
        Dim varIssuedQty As Object = Nothing
        Dim varRcvdQty As Object = Nothing
        Dim varRevNo As Object = Nothing
        Dim rsbagqty As SqlCommand
        Dim rdrbagQty As SqlDataReader
        Dim dailypullrate As Long
        Dim intPos As Integer
        Dim sqlInsertUpdate As SqlCommand
        Dim CONSIGNEE_CODE As String
        Dim varWHCODE As Object = Nothing
        Dim blnDAILYPULLFLAG As Integer
        Dim Rs As SqlCommand
        Dim rdr As SqlDataReader
        Dim sqltrans As SqlTransaction
        Dim rsTransitDays As ADODB.Recordset
        Dim WHDATE As Date
        Dim dtDate As Date
        Dim lngBackDateSch As Long
        Dim dtPrevDate As Date

        blnDAILYPULLFLAG = 0

        sqlCMDDisplay = New SqlCommand
        sqlCMDDisplay.Connection = SqlConnectionclass.GetConnection

        sqltrans = sqlCMDDisplay.Connection.BeginTransaction
        sqlCMDDisplay.Transaction = sqltrans

        sqlCMDDisplay.CommandText = "DELETE FROM TMP_VDAPROPOSAL WHERE UNIT_CODE = '" & gstrUNITID & "'"
        sqlCMDDisplay.ExecuteNonQuery()

        sqlCMDDisplay.CommandText = "INSERT INTO TMP_VDAPROPOSAL SELECT * FROM VW_SCHEDULE_PROPOSAL WHERE DOC_NO = " & txtDocNo.Text & " AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCMDDisplay.ExecuteNonQuery()

        If FileType = "VDA" Then
            sql = " Select Distinct VW.DDRD_Req_Dt1,VW.Cust_Drgno," & _
                  " VW.Item_Code,VW.Item_Desc,VW.GI_Vend_Code, " & _
                  " VW.SAFETYSTKPERDAY,SP.SAFETYDAYSMAX,SP.SAFETYDAYS,SP.SAFETYDAYSMIN," & _
                  " VW.DDRD_Req_Qty1 AS SHIPQTY,vw.FactoryCode AS FACTORY_CODE,vw.CONSIGNEE_CODE," & _
                  " VW.StockCalcWAdays, VW.ScheduleCalcMonths,VW.DAYSFORSAFETYSTOCK,VW.SUMOFRELEASEQTY  " & _
                  " From tmp_Schedule_Uploading_Darwin VW with (nolock), SCHEDULEPARAMETER_DTL SP with (nolock)," & _
                  " Custitem_mst CM with (nolock),customer_mst CMS with (nolock)"
            sql = sql + " Where CM.Account_code=SP.Consignee_code AND CM.UNIT_CODE = SP.UNIT_CODE"
            sql = sql + " and CM.Cust_drgno= VW.Cust_Drgno and CM.Active=1 AND CM.SCHUPLDREQD = 1 AND CM.UNIT_CODE = VW.UNIT_CODE"
            sql = sql + " and VW.Item_code=CM.Item_code AND VW.UNIT_CODE = CM.UNIT_CODE"
            sql = sql + " And VW.Doc_No=" & TRANS_NUMBER & ""
            sql = sql + " AND SP.CUSTOMER_CODE  = '" & txtCustomerCode.Text & "'  "
            sql = sql + " AND SP.CUST_DRGNO = VW.Cust_Drgno AND SP.UNIT_CODE = VW.UNIT_CODE"
            sql = sql + " AND SP.WH_CODE = VW.GI_Vend_Code "
            sql = sql + " and CMS.Cust_Vendor_Code = VW.GI_Vend_Code AND CMS.UNIT_CODE = VW.UNIT_CODE"
            sql = sql + " AND CMS.DOCK_CODE = VW.FACTORYCODE AND SP.CONSIGNEE_CODE = VW.CONSIGNEE_CODE AND VW.UNIT_CODE = '" & gstrUNITID & "' "

            If Len(strWHCode) > 0 Then sql = sql & " AND VW.GI_Vend_Code not in (" & strWHCode & ") "
            sql = sql + " Order By VW.Cust_Drgno,VW.Item_Code,VW.DDRD_Req_Dt1 "
        End If

        If FileType = "EDIFACT" Then
            '10902693
            sql = " Select Distinct D.Delivery_Dt,C.Cust_Drgno,I.Item_Code, I.DESCRIPTION,H.PARTY_ID1,H.PARTY_ID3,T.SAFETYSTOCK AS SAFETYSTKPERDAY, SP.SAFETYDAYSMAX, " & _
                    " SP.SAFETYDAYS,SP.Consignee_code ,SP.SAFETYDAYSMIN,sum(QUANTITY) AS SHIPQTY, D.FREQUENCY,D.Dispatch_Pattern, T.StockCalcWAdays,T.ScheduleCalcMonths," & _
                    " T.DAYSFORSAFETYSTOCK,T.SUMOFRELEASEQTY   From SCHEDULE_UPLOAD_DARWINEDIFACT_DTL D with (nolock), CUSTITEM_MST C with (nolock),ITEM_MST I with (nolock), " & _
                    " TMPSCHEDULESAFETYSTOCK T with (nolock),  SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H with (nolock), SCHEDULEPARAMETER_DTL SP with (nolock) " & _
                    " Where C.Account_code = SP.Consignee_code And C.UNIT_CODE = SP.UNIT_CODE And C.Cust_drgno = T.CUSTDRG_NO And C.UNIT_CODE = T.UNIT_CODE And C.Active = 1" & _
                    " AND C.SCHUPLDREQD = 1 And H.Doc_No= " & TRANS_NUMBER & " AND D.ITEM_CODE = C.CUST_DRGNO AND D.UNIT_CODE = C.UNIT_CODE AND C.ITEM_CODE = I.ITEM_CODE " & _
                    " AND C.UNIT_CODE = I.UNIT_CODE AND T.CUSTDRG_NO = C.CUST_DRGNO AND T.UNIT_CODE = C.UNIT_CODE AND D.DOC_NO = H.DOC_NO AND D.UNIT_CODE = H.UNIT_CODE" & _
                    " AND SP.CUSTOMER_CODE  = '" & txtCustomerCode.Text & "' AND D.UNIT_CODE = '" & gstrUNITID & "'  AND SP.CUST_DRGNO = T.CUSTDRG_NO AND SP.UNIT_CODE = T.UNIT_CODE   " & _
                    " AND SP.WH_CODE = H.PARTY_ID1 AND SP.UNIT_CODE = H.UNIT_CODE  AND T.WH_CODE = H.PARTY_ID1 AND T.UNIT_CODE = H.UNIT_CODE   " & _
                    " group by D.Delivery_Dt,C.Cust_Drgno,I.Item_Code, I.DESCRIPTION,H.PARTY_ID1,H.PARTY_ID3,T.SAFETYSTOCK , SP.SAFETYDAYSMAX, SP.SAFETYDAYS,SP.Consignee_code ," & _
                    " SP.SAFETYDAYSMIN, D.FREQUENCY, D.Dispatch_Pattern, T.StockCalcWAdays, T.ScheduleCalcMonths, T.DAYSFORSAFETYSTOCK, T.SUMOFRELEASEQTY  Order By D.DELIVERY_DT "
        End If

        If FileType = "COVISINT" Then
            sql = "Select Distinct D.Delivery_DATE,C.Cust_Drgno,I.Item_Code, " & _
                " I.DESCRIPTION,D.WH_CODE,T.SAFETYSTOCK AS SAFETYSTKPERDAY, " & _
                " SP.SAFETYDAYSMAX,SP.SAFETYDAYS  ,SP.SAFETYDAYSMIN,D.QTY AS SHIPQTY, " & _
                " D.FACTORY_CODE, D.CONSIGNEE_CODE, T.StockCalcWAdays, T.ScheduleCalcMonths, T.DAYSFORSAFETYSTOCK, T.SUMOFRELEASEQTY" & _
                " From SCHEDULE_UPLOAD_COVISINT_DTL D with (nolock), CUSTITEM_MST C with (nolock),ITEM_MST I with (nolock), TMPSCHEDULESAFETYSTOCK T with (nolock)," & _
                " SCHEDULE_UPLOAD_COVISINT_HDR H with (nolock), SCHEDULEPARAMETER_DTL SP with (nolock), customer_mst cm" & _
                " Where(C.Account_code = SP.Consignee_code AND C.UNIT_CODE = SP.UNIT_CODE And D.CONSIGNEE_code = sp.CONSIGNEE_code AND D.UNIT_CODE = SP.UNIT_CODE)" & _
                " and D.FACTORY_CODE = cm.dock_code AND D.UNIT_CODE = CM.UNIT_CODE" & _
                " and t.consignee_code = d.consignee_code AND T.UNIT_CODE = D.UNIT_CODE" & _
                " and cm.customer_code = d.consignee_code AND CM.UNIT_CODE = D.UNIT_CODE" & _
                " and C.Cust_drgno= T.CUSTDRG_NO AND C.UNIT_CODE = T.UNIT_CODE and C.Active=1 AND C.SCHUPLDREQD = 1" & _
                " And H.Doc_No= " & TRANS_NUMBER & " " & _
                " AND D.ITEM_CODE = C.CUST_DRGNO AND C.UNIT_CODE = D.UNIT_CODE AND C.ITEM_CODE = I.ITEM_CODE AND C.UNIT_CODE = I.UNIT_CODE" & _
                " AND T.CUSTDRG_NO = C.CUST_DRGNO " & _
                " AND D.DOC_NO = H.DOC_NO AND D.UNIT_CODE = H.UNIT_CODE  " & _
                " AND SP.CUSTOMER_CODE  = '" & Me.txtCustomerCode.Text & "'  " & _
                " AND SP.CUST_DRGNO = T.CUSTDRG_NO AND SP.UNIT_CODE = T.UNIT_CODE  " & _
                " AND SP.WH_CODE = D.WH_CODE  AND SP.UNIT_CODE = D.UNIT_CODE" & _
                " AND T.WH_CODE = SP.WH_CODE AND T.UNIT_CODE = SP.UNIT_CODE AND D.UNIT_CODE = '" & gstrUNITID & "'" & _
                " Order By C.Cust_Drgno,D.DELIVERY_DATE "
        End If

        sqlCMDDisplay.CommandText = sql
        sqlRdrDisplay = sqlCMDDisplay.ExecuteReader

        SCHQTY = 0

        Dim COUNT As Integer
        COUNT = 0
        If sqlRdrDisplay.HasRows Then
            rsbagqty = New SqlCommand
            rsbagqty.Connection = SqlConnectionclass.GetConnection
            Rs = New SqlCommand
            Rs.Connection = SqlConnectionclass.GetConnection

            While sqlRdrDisplay.Read

                rsbagqty.CommandText = "select bag_qty from item_mst where item_code = '" & sqlRdrDisplay("Item_Code").ToString & "' and Status = 'A' AND UNIT_CODE = '" & gstrUNITID & "'"
                rdrbagQty = rsbagqty.ExecuteReader
                If rdrbagQty.HasRows Then
                    rdrbagQty.Read()
                    mlngBAGQTY = rdrbagQty("BAG_QTY").ToString
                Else
                    mlngBAGQTY = 1
                End If
                rdrbagQty.Close()

                sql = " Select TransitDaysBySea, BufferDays "
                sql = sql & "  From ScheduleParameter_mst"
                sql = sql & " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and consignee_code = '" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "' AND UNIT_CODE = '" & gstrUNITID & "'"

                rsbagqty.CommandText = sql
                rdrbagQty = rsbagqty.ExecuteReader

                If rdrbagQty.HasRows Then
                    rdrbagQty.Read()
                    Transit_Days = IIf(IsDBNull(rdrbagQty("TransitDaysBySea").ToString), 0, rdrbagQty("TransitDaysBySea").ToString)
                    Buffer_Days = IIf(IsDBNull(rdrbagQty("BufferDays").ToString), 0, rdrbagQty("BufferDays").ToString)
                End If
                rdrbagQty.Close()

                If FileType = "EDIFACT" Then
                    If sqlRdrDisplay("Frequency").ToString = "" And sqlRdrDisplay("DISPATCH_PATTERN").ToString = "" Then
                        GoTo SKIP
                    End If
                End If

                rsDate = New SqlCommand
                rsDate.Connection = SqlConnectionclass.GetConnection

                sql = "set dateformat 'dmy' select max(dt) as dt from Calendar_Mfg_mst" & _
                    " where work_flg=0 AND UNIT_CODE = '" & gstrUNITID & "' and "

                If FileType = "VDA" Then
                    sql = sql + " dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), VB6.Format(CDate(sqlRdrDisplay("DDRD_Req_Dt1").ToString), "dd MMM yyyy")) & "'  "
                End If

                If FileType = "EDIFACT" Then
                    sql = sql + " dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), VB6.Format(CDate(sqlRdrDisplay("DELIVERY_DT").ToString), "dd MMM yyyy")) & "' "
                End If

                If FileType = "COVISINT" Then
                    sql = sql + " dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), VB6.Format(CDate(sqlRdrDisplay("DELIVERY_DATE").ToString), "dd MMM yyyy")) & "' "
                End If

                rsDate.CommandText = sql
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                rdr = rsDate.ExecuteReader
                If rdr.HasRows Then
                    rdr.Read()
                    dtDate = CDate(IIf(rdr("dt").ToString = "", "01 jan 1900", rdr("dt").ToString))
                End If
                rdr.Close()

                If CDate(getDateForDB(dtDate)) >= CDate(getDateForDB(GetServerDate())) Then

                    sql = "SELECT WHSTOCK,MINSTOCK,RCVDQTY,ISSUEDQTY,WHDATE,REVNO FROM TMPWHSTOCK " & _
                         " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                         " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "

                    If FileType = "VDA" Then
                        sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("GI_VEND_CODE").ToString & "' " & _
                            " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"
                        varWHCODE = sqlRdrDisplay("GI_VEND_CODE").ToString
                    End If

                    If FileType = "EDIFACT" Then
                        sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("PARTY_ID1").ToString & "' "
                        varWHCODE = sqlRdrDisplay("PARTY_ID1").ToString
                    End If

                    If FileType = "COVISINT" Then
                        sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("wh_code").ToString & "' " & _
                            " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"
                        varWHCODE = sqlRdrDisplay("wh_code").ToString
                    End If

                    RSWHSTOCK.GetResult(sql, ADODB.CursorTypeEnum.adOpenStatic)

                    If RSWHSTOCK.GetNoRows > 0 Then
                        varWhStock = RSWHSTOCK.GetValue("WHSTOCK")
                        varIssuedQty = RSWHSTOCK.GetValue("ISSUEDQTY")
                        varRcvdQty = RSWHSTOCK.GetValue("RCVDQTY")
                        varRevNo = RSWHSTOCK.GetValue("REVNO")
                    Else
                        varWhStock = 0
                        varIssuedQty = 0
                        varRcvdQty = 0
                        varRevNo = 0
                    End If

                    If varWhStock < 0 Then
                        If FileType = "VDA" Then
                            sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                                " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                                " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                                " Updt_Dt,Updt_Uid,WH_DATE,REVNO,CALLOFFNORESETREMARKS,DAILYPULLFLAG" & _
                                " ,SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                                " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays, bufferDays,UNIT_CODE)" & _
                                " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("GI_VEND_CODE").ToString & "', " & _
                                " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "',CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                " 0,CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                " '" & -(varWhStock) & "','" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "','" & varWhStock & "'," & _
                                " 0,'" & -(varWhStock) & "',getDate(),'" & mP_User & "',getDate(),'" & mP_User & "'," & _
                                " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "' , '" & blnDAILYPULLFLAG & "'," & _
                                " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                                " " & sqlRdrDisplay("SafetyDays").ToString & "," & mlngBAGQTY & "," & _
                                " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                                " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                                " " & Buffer_Days & ",'" & gstrUNITID & "')"

                            sqlInsertUpdate.CommandText = sql
                            sqlInsertUpdate.ExecuteNonQuery()

                            sql = "UPDATE TMPWHSTOCK SET WHSTOCK = 0 " & _
                                " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                                " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' " & _
                                " AND WH_CODE = '" & sqlRdrDisplay("GI_VEND_CODE").ToString & "' " & _
                                " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"

                            sqlInsertUpdate.CommandText = sql
                            sqlInsertUpdate.ExecuteNonQuery()

                        End If

                        If FileType = "EDIFACT" Then
                            sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                                    " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                                    " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                                    " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
                                    " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                                    " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
                                    " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("PARTY_ID1").ToString & "', " & _
                                    " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "',CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                    " 0,CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                    " '" & -(varWhStock) & "','" & CONSIGNEE_CODE & "','" & varWhStock & "'," & _
                                    " 0,'" & -(varWhStock) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                                    " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                                    " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                                    " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                                    " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                                    " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                                    " " & Buffer_Days & ",'" & gstrUNITID & "' )"

                            sqlInsertUpdate.CommandText = sql
                            sqlInsertUpdate.ExecuteNonQuery()

                            sql = "UPDATE TMPWHSTOCK SET WHSTOCK = 0 " & _
                                " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                                " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' " & _
                                " AND WH_CODE = '" & sqlRdrDisplay("PARTY_ID1").ToString & "' "

                            sqlInsertUpdate.CommandText = sql
                            sqlInsertUpdate.ExecuteNonQuery()
                        End If

                        If FileType = "COVISINT" Then
                            sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                                " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                                " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                                " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG" & _
                                " ,SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                                " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY, transitDays,bufferDays,UNIT_CODE)" & _
                                " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("wh_code").ToString & "', " & _
                                " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "',Convert(DateTime, Convert(varchar(11), getdate(), 106), 106) ," & _
                                " 0,Convert(DateTime, Convert(varchar(11), getdate(), 106), 106)," & _
                                " '" & -(varWhStock) & "','" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "','" & varWhStock & "'," & _
                                " 0,'" & -(varWhStock) & "',getDate(),'" & mP_User & "' ,getDate(),'" & mP_User & "'," & _
                                " '" & Format(WHDATE, "dd MMM yyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                                " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                                " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                                " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                                " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                                " " & Buffer_Days & ",'" & gstrUNITID & "' )"

                            sqlInsertUpdate.CommandText = sql
                            sqlInsertUpdate.ExecuteNonQuery()

                            sql = "UPDATE TMPWHSTOCK SET WHSTOCK = 0 " & _
                                " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                                " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' " & _
                                " AND WH_CODE = '" & sqlRdrDisplay("wh_code").ToString & "' " & _
                                " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"

                            sqlInsertUpdate.CommandText = sql
                            sqlInsertUpdate.ExecuteNonQuery()
                        End If
                    End If
                End If

                With Rs
                    .CommandType = CommandType.StoredProcedure
                    If FileType = "VDA" Then
                        .CommandText = "scheduleQty"
                        .CommandTimeout = 0
                    End If
                    If FileType = "EDIFACT" Then
                        .CommandText = "SCHEDULEQTY_EDIFACT_MTL"
                        .CommandTimeout = 0
                    End If
                    If FileType = "COVISINT" Then
                        .CommandText = "SCHEDULEQTY_COVISINT"
                        .CommandTimeout = 0
                    End If
                    .Parameters.Clear()
                    .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 8, Trim(txtCustomerCode.Text)).Direction = ParameterDirection.Input
                    .Parameters.Add("@CUST_CODE", SqlDbType.VarChar, 8, Trim(txtCustomerCode.Text)).Direction = ParameterDirection.Input
                    .Parameters.Add("@DOCNO", SqlDbType.Int, 12, Trim(TRANS_NUMBER)).Direction = ParameterDirection.Input
                    .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 50, Trim(sqlRdrDisplay("Cust_DrgNo").ToString)).Direction = ParameterDirection.Input

                    If FileType = "VDA" Then
                        .Parameters.Add("@WH_CODE", SqlDbType.VarChar, 20, sqlRdrDisplay("GI_VEND_CODE").ToString).Direction = ParameterDirection.Input
                        .Parameters.Add("@TRANS_DT", SqlDbType.Date, 11, Format(CDate(sqlRdrDisplay("DDRD_Req_Dt1")), "dd MMM yyyy")).Direction = ParameterDirection.Input
                        .Parameters.Add("@FACTORYCODE", SqlDbType.VarChar, 10, Trim(sqlRdrDisplay("Factory_Code").ToString)).Direction = ParameterDirection.Input
                        .Parameters.Add("@CONSIGNEE_CODE", SqlDbType.VarChar, 8, Trim(sqlRdrDisplay("CONSIGNEE_CODE").ToString)).Direction = ParameterDirection.Input

                        .Parameters(0).Value = Trim(gstrUNITID) : .Parameters(1).Value = Trim(txtCustomerCode.Text) : .Parameters(2).Value = Trim(TRANS_NUMBER)
                        .Parameters(3).Value = Trim(sqlRdrDisplay("Cust_DrgNo").ToString) : .Parameters(4).Value = sqlRdrDisplay("GI_VEND_CODE").ToString
                        .Parameters(5).Value = Format(CDate(sqlRdrDisplay("DDRD_Req_Dt1")), "dd MMM yyyy") : .Parameters(6).Value = Trim(sqlRdrDisplay("Factory_Code").ToString)
                        .Parameters(7).Value = Trim(sqlRdrDisplay("CONSIGNEE_CODE").ToString)
                    End If

                    If FileType = "EDIFACT" Then
                        .Parameters.Add("@WH_CODE", SqlDbType.VarChar, 12, sqlRdrDisplay("PARTY_ID1").ToString).Direction = ParameterDirection.Input
                        .Parameters.Add("@TRANS_DT", SqlDbType.Date, 11, Format(CDate(Trim(sqlRdrDisplay("DELIVERY_DT"))), "dd MMM yyyy")).Direction = ParameterDirection.Input
                        .Parameters.Add("@FACTORYCODE", SqlDbType.VarChar, 10, sqlRdrDisplay("PARTY_ID3").ToString).Direction = ParameterDirection.Input

                        rsbagqty.CommandText = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST" & _
                            " WHERE CUST_VENDOR_CODE = '" & sqlRdrDisplay("PARTY_ID1").ToString & "'" & _
                            " AND DOCK_CODE = '" & sqlRdrDisplay("PARTY_ID3").ToString & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                        rdrbagQty = rsbagqty.ExecuteReader
                        If rdrbagQty.HasRows Then
                            rdrbagQty.Read()
                            CONSIGNEE_CODE = Trim(rdrbagQty("CUSTOMER_CODE").ToString)
                        Else
                            CONSIGNEE_CODE = ""
                        End If

                        rdrbagQty.Close()

                        .Parameters(0).Value = Trim(gstrUNITID) : .Parameters(1).Value = Trim(txtCustomerCode.Text) : .Parameters(2).Value = Trim(TRANS_NUMBER)
                        .Parameters(3).Value = Trim(sqlRdrDisplay("Cust_DrgNo").ToString) : .Parameters(4).Value = sqlRdrDisplay("PARTY_ID1").ToString
                        .Parameters(5).Value = Format(CDate(Trim(sqlRdrDisplay("DELIVERY_DT"))), "dd MMM yyyy") : .Parameters(6).Value = sqlRdrDisplay("PARTY_ID3").ToString
                        ' .Parameters(7).Value = CONSIGNEE_CODE

                    End If

                    If FileType = "COVISINT" Then
                        .Parameters.Add("@WH_CODE", SqlDbType.VarChar, 12, sqlRdrDisplay("wh_code").ToString).Direction = ParameterDirection.Input
                        .Parameters.Add("@TRANS_DT", SqlDbType.Date, 11, Trim(sqlRdrDisplay("DELIVERY_DATE"))).Direction = ParameterDirection.Input
                        .Parameters.Add("@FACTORYCODE", SqlDbType.VarChar, 10, Trim(sqlRdrDisplay("Factory_Code").ToString)).Direction = ParameterDirection.Input
                        .Parameters.Add("@CONSIGNEE_CODE", SqlDbType.VarChar, 10, Trim(sqlRdrDisplay("CONSIGNEE_CODE").ToString)).Direction = ParameterDirection.Input

                        .Parameters(0).Value = Trim(gstrUNITID) : .Parameters(1).Value = Trim(txtCustomerCode.Text) : .Parameters(2).Value = Trim(TRANS_NUMBER)
                        .Parameters(3).Value = Trim(sqlRdrDisplay("Cust_DrgNo").ToString) : .Parameters(4).Value = sqlRdrDisplay("wh_code").ToString
                        .Parameters(5).Value = Format(CDate(sqlRdrDisplay("DELIVERY_DATE")), "dd MMM yyyy") : .Parameters(6).Value = Trim(sqlRdrDisplay("Factory_Code").ToString)
                        .Parameters(7).Value = Trim(sqlRdrDisplay("CONSIGNEE_CODE").ToString)
                    End If

                    .Parameters.Add("@SCHQTY", SqlDbType.Int, 12, 0).Direction = ParameterDirection.Output

                    Rs.ExecuteScalar()

                End With

                sql = "SELECT WHSTOCK,MINSTOCK,RCVDQTY,ISSUEDQTY,WHDATE,REVNO FROM TMPWHSTOCK " & _
                         " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                         " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "

                If FileType = "VDA" Then
                    sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("GI_VEND_CODE").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"
                    varWHCODE = sqlRdrDisplay("GI_VEND_CODE").ToString
                End If

                If FileType = "EDIFACT" Then
                    sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("PARTY_ID1").ToString & "' "
                    varWHCODE = sqlRdrDisplay("PARTY_ID1").ToString
                End If

                If FileType = "COVISINT" Then
                    sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("wh_code").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"
                    varWHCODE = sqlRdrDisplay("wh_code").ToString
                End If

                RSWHSTOCK.GetResult(sql, ADODB.CursorTypeEnum.adOpenStatic)

                If RSWHSTOCK.GetNoRows > 0 Then
                    varWhStock = RSWHSTOCK.GetValue("WHSTOCK")
                    varIssuedQty = RSWHSTOCK.GetValue("ISSUEDQTY")
                    varRcvdQty = RSWHSTOCK.GetValue("RCVDQTY")
                    varRevNo = RSWHSTOCK.GetValue("REVNO")
                Else
                    varWhStock = 0
                    varIssuedQty = 0
                    varRcvdQty = 0
                    varRevNo = 0
                End If

                If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(VB6.Format(dtDate, "dd MMM yyyy")), CDate(VB6.Format(GetServerDate(), "dd MMM yyyy"))) > 0 Then
                    SCHQTY = 0
                    GoTo NOSAFETYSTOCKPERDAY
                End If

                If RSWHSTOCK.GetNoRows > 0 Then
                    If Val(varWhStock) + Val(varRcvdQty) > Val(varIssuedQty) Then
                        SCHQTY = varIssuedQty
                    Else
                        If Val(varWhStock) > 0 Then
                            SCHQTY = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)
                        Else
                            SCHQTY = Val(varIssuedQty) - (Val(varWhStock) + Val(varRcvdQty))
                        End If
                    End If
                Else
                    SCHQTY = IIf(IsDBNull(Rs.Parameters("@SCHQTY").Value), 0, Rs.Parameters("@SCHQTY").Value)
                End If

                If chkDlyPullQty.Checked = True Then

                    Dim schqty1 As String

                    schqty1 = FN_DAILYPULLQTY(varWhStock, varIssuedQty, varRcvdQty, SCHQTY, varWHCODE, sqlRdrDisplay("SAFETYDAYS").ToString, CStr(sqlRdrDisplay("Item_Code").ToString), CStr(sqlRdrDisplay("Cust_DrgNo").ToString), CInt(Row), CStr(sqlRdrDisplay("Factory_Code").ToString), mlngBAGQTY)
                    intPos = InStr(1, schqty1, "*")
                    dailypullrate = 0
                    If intPos > 0 Then
                        dailypullrate = Mid(schqty1, intPos + 1, Len(schqty1))
                        SCHQTY = Mid(schqty1, 1, intPos - 1)
                    End If
                    blnDAILYPULLFLAG = 1
                Else
                    If Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) > 0 Then
                        If ((Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) * Val(sqlRdrDisplay("safetydaysmin").ToString)) - SCHQTY) <= 0 Then
                            SCHQTY = 0
                        Else
                            SAFETYDAYS_BELOW = (Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) * Val(sqlRdrDisplay("safetydaysmin").ToString)) - SCHQTY ''/ val(adors!SAFETYSTKPERDAY)
                            SFTYDAYS_MNTD = Val(sqlRdrDisplay("safetydaysMAX").ToString) - Val(sqlRdrDisplay("safetydaysmin").ToString)  ''+ SAFETYDAYS_BELOW
                            SFTYDAYS_MNTD = System.Math.Round(SFTYDAYS_MNTD, 0)

                            SCHQTY = (Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) * SFTYDAYS_MNTD) + SAFETYDAYS_BELOW

                            SCHQTY = System.Math.Round(SCHQTY, 0)

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
                    Else
                        If SCHQTY > 0 Then
                            If (Val(varRcvdQty) + Val(varWhStock)) >= SCHQTY Then
                                SCHQTY = 0
                                GoTo NOSAFETYSTOCKPERDAY
                            Else
                                SCHQTY = Val(varIssuedQty) - (Val(varRcvdQty) + Val(varWhStock))
                            End If
                        Else
                            SCHQTY = Val(varIssuedQty) - (Val(varRcvdQty) + Val(varWhStock))
                        End If

                        SFTYDAYS_MNTD = Val(sqlRdrDisplay("SafetyDaysMax").ToString) - Val(sqlRdrDisplay("safetydaysmin").ToString) ''+ SAFETYDAYS_BELOW
                        SFTYDAYS_MNTD = System.Math.Round(SFTYDAYS_MNTD, 0)

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
                        Else
                            SCHQTY = 0
                        End If

                    End If
                End If

NOSAFETYSTOCKPERDAY:
                sql = "select top 1 trans_dt,customer_code," & _
                    " revno from warehouse_stock_dtl" & _
                    " where customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"

                If FileType = "VDA" Then
                    sql = sql & " and  warehouse_code = '" & sqlRdrDisplay("GI_VEND_CODE").ToString & "'"
                End If

                If FileType = "EDIFACT" Then
                    sql = sql & " and  warehouse_code = '" & sqlRdrDisplay("PARTY_ID1").ToString & "'"
                End If

                If FileType = "COVISINT" Then
                    sql = sql + " and  warehouse_code = '" & sqlRdrDisplay("wh_code").ToString & "'"
                End If

                sql = sql + " group by customer_code, trans_dt,revno" & _
                " order by trans_dt desc,revno desc "

                rsWHDt = New SqlCommand
                rsWHDt.Connection = SqlConnectionclass.GetConnection
                rsWHDt.CommandText = sql
                rdrWHDt = rsWHDt.ExecuteReader

                If rdrWHDt.HasRows Then
                    rdrWHDt.Read()
                    WHDATE = rdrWHDt("TRANS_DT").ToString
                Else
                    WHDATE = ""
                End If

                rdrWHDt.Close()
                rsWHDt.Dispose()
                rsWHDt = Nothing

                If chkDlyPullQty.Checked = True Then
                    blnDAILYPULLFLAG = 1
                Else
                    blnDAILYPULLFLAG = 0
                End If

                If Val(SCHQTY) > 0 Then
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

                If sqlInsertUpdate Is Nothing Then
                    sqlInsertUpdate = New SqlCommand
                    sqlInsertUpdate.Connection = SqlConnectionclass.GetConnection
                End If

                If FileType = "VDA" Then

                    sqlInsertUpdate.CommandText = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                        " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                        " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                        " Updt_Dt,Updt_Uid,WH_DATE,REVNO,CALLOFFNORESETREMARKS,DAILYPULLFLAG" & _
                        " ,SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                        " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays, bufferDays,UNIT_CODE)" & _
                        " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("GI_VEND_CODE").ToString & "', " & _
                        " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "','" & Format(CDate(sqlRdrDisplay("DDRD_Req_Dt1")), "dd MMM yyyy") & "'," & _
                        " '" & Val(sqlRdrDisplay("shipqty").ToString) & "','" & getDateForDB(dtDate) & "'," & _
                        " '" & SCHQTY & "','" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "','" & varWhStock & "'," & _
                        " '" & varRcvdQty & "','" & varIssuedQty & "',getDate(),'" & mP_User & "',getDate(),'" & mP_User & "'," & _
                        " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "' , '" & blnDAILYPULLFLAG & "'," & _
                        " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                        " " & sqlRdrDisplay("SafetyDays").ToString & "," & mlngBAGQTY & "," & _
                        " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                        " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                        " " & Buffer_Days & ",'" & gstrUNITID & "')"
                    sqlInsertUpdate.ExecuteNonQuery()
                End If

                If FileType = "EDIFACT" Then
                    sqlInsertUpdate.CommandText = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                            " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                            " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                            " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
                            " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                            " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
                            " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("PARTY_ID1").ToString & "', " & _
                            " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "','" & Format(CDate(sqlRdrDisplay("DELIVERY_DT").ToString), "dd MMM yyyy") & "'," & _
                            " '" & Val(sqlRdrDisplay("shipqty").ToString) & "','" & Format(dtDate, "dd MMM yyyy") & "'," & _
                            " '" & SCHQTY & "','" & CONSIGNEE_CODE & "','" & IIf(IsDBNull(varWhStock), 0, varWhStock) & "'," & _
                            " '" & varRcvdQty & "','" & varIssuedQty & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                            " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                            " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                            " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                            " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                            " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                            " " & Buffer_Days & ",'" & gstrUNITID & "' )"
                    sqlInsertUpdate.ExecuteNonQuery()
                End If

                If FileType = "COVISINT" Then
                    sqlInsertUpdate.CommandText = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                        " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                        " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                        " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG" & _
                        " ,SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                        " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY, transitDays,bufferDays,UNIT_CODE)" & _
                        " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("wh_code").ToString & "', " & _
                        " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "','" & Format(CDate(sqlRdrDisplay("DELIVERY_DATE").ToString), "dd MMM yyyy") & "'," & _
                        " '" & Val(sqlRdrDisplay("shipqty").ToString) & "','" & Format(dtDate, "dd MMM yyy") & "'," & _
                        " '" & SCHQTY & "','" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "','" & varWhStock & "'," & _
                        " '" & varRcvdQty & "','" & varIssuedQty & "',getDate(),'" & mP_User & "' ,getDate(),'" & mP_User & "'," & _
                        " '" & Format(WHDATE, "dd MMM yyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                        " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                        " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                        " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                        " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                        " " & Buffer_Days & ",'" & gstrUNITID & "' )"
                    sqlInsertUpdate.ExecuteNonQuery()
                End If

                TMPWHSTOCK = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)

                If Val(sqlRdrDisplay("shipqty").ToString) >= 0 Then
                    TMPWHSTOCK = TMPWHSTOCK + SCHQTY                        ''- val(adors!SHIPQTY)
                End If
                updSQL = "UPDATE TMPWHSTOCK" & _
                    " SET WHSTOCK = " & TMPWHSTOCK & " , "

                If FileType = "VDA" Then
                    updSQL = updSQL + " WHDATE = '" & VB6.Format(sqlRdrDisplay("DDRD_Req_Dt1"), "dd MMM yyyy") & "'"
                End If

                If FileType = "EDIFACT" Then
                    updSQL = updSQL + " WHDATE = '" & VB6.Format(sqlRdrDisplay("DELIVERY_DT"), "dd MMM yyyy") & "'"
                End If

                If FileType = "COVISINT" Then
                    updSQL = updSQL + " WHDATE = '" & VB6.Format(sqlRdrDisplay("DELIVERY_DATE"), "dd MMM yyyy") & "'"
                End If

                updSQL = updSQL + " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                    " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "

                If FileType = "VDA" Then
                    updSQL = updSQL + " AND WH_CODE = '" & sqlRdrDisplay("GI_VEND_CODE").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"
                End If

                If FileType = "EDIFACT" Then
                    updSQL = updSQL + " AND WH_CODE = '" & sqlRdrDisplay("PARTY_ID1").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("PARTY_ID3").ToString) & "'"
                End If

                If FileType = "COVISINT" Then
                    updSQL = updSQL + " AND WH_CODE = '" & sqlRdrDisplay("wh_code").ToString & "' " & _
                        " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("Factory_Code").ToString) & "'"
                End If

                mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(updSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

SKIP:

            End While

            sqlRdrDisplay.Close()

            If Not sqlInsertUpdate Is Nothing Then
                sqlInsertUpdate.Dispose()
                sqlInsertUpdate = Nothing
            End If
        Else
            sqlRdrDisplay.Close()
        End If

        If Not sqlCMDDisplay Is Nothing Then
            sqlCMDDisplay.Dispose()
            sqlCMDDisplay = Nothing
        End If

        If Not Rs Is Nothing Then
            Rs.Dispose()
            Rs = Nothing
        End If

        If Not rsDate Is Nothing Then
            rsDate.Dispose()
            rsDate = Nothing
        End If

        If Not sqltrans Is Nothing Then
            sqltrans.Commit()
            sqltrans = Nothing
        End If

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Function

ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        If Not sqlCMDDisplay Is Nothing Then
            sqlCMDDisplay.Dispose()
            sqlCMDDisplay = Nothing
        End If

        If Not Rs Is Nothing Then
            Rs.Dispose()
            Rs = Nothing
        End If

        RSWHSTOCK.ResultSetClose() : RSWHSTOCK = Nothing

        If Not rsDate Is Nothing Then
            rsDate.Dispose()
            rsDate = Nothing
        End If

        If Not sqltrans Is Nothing Then
            sqltrans.Rollback()
            sqltrans = Nothing
        End If

        Return ""
        Exit Function

    End Function

    Private Function FN_Display_PKC(ByVal TRANS_NUMBER As String, ByVal FileType As String) As String
        On Error GoTo ERR_Renamed

        Dim Transit_Days, Row, Buffer_Days As Integer
        Dim SCHQTY As Long
        Dim sqlCMDDisplay As SqlCommand
        Dim sqlRdrDisplay As SqlDataReader
        Dim SFTYDAYS_MNTD As Object
        Dim SAFETYDAYS_BELOW As Long
        Dim TMPWHSTOCK As Long
        Dim rsDate As SqlCommand
        Dim sql As String = String.Empty, updSQL As String = String.Empty, strWHCode As String = String.Empty
        Dim RSWHSTOCK As New ClsResultSetDB
        Dim rsWHDt As SqlCommand
        Dim rdrWHDt As SqlDataReader
        Dim varWhStock As Object = Nothing
        Dim varIssuedQty As Object = Nothing
        Dim varRcvdQty As Object = Nothing
        Dim varRevNo As Object = Nothing
        Dim rsbagqty As SqlCommand
        Dim rdrbagQty As SqlDataReader
        Dim dailypullrate As Long
        Dim intPos As Integer
        Dim sqlInsertUpdate As SqlCommand
        Dim CONSIGNEE_CODE As String
        Dim varWHCODE As Object = Nothing
        Dim blnDAILYPULLFLAG As Integer
        Dim Rs As SqlCommand
        Dim rdr As SqlDataReader
        Dim sqltrans As SqlTransaction
        Dim rsTransitDays As ADODB.Recordset
        Dim WHDATE As Date
        Dim dtDate As Date
        Dim lngBackDateSch As Long
        Dim dtPrevDate As Date

        blnDAILYPULLFLAG = 0

        sqlCMDDisplay = New SqlCommand
        sqlCMDDisplay.Connection = SqlConnectionclass.GetConnection

        sqltrans = sqlCMDDisplay.Connection.BeginTransaction
        sqlCMDDisplay.Transaction = sqltrans

        sql = "  Select Distinct D.ScheduleDate,C.Cust_Drgno,I.Item_Code, I.DESCRIPTION,H.suppliercode,T.SAFETYSTOCK AS SAFETYSTKPERDAY," & _
            " SP.SAFETYDAYSMAX, SP.SAFETYDAYS,SP.Consignee_code ,SP.SAFETYDAYSMIN,sum(D.DispatchQty) AS SHIPQTY,T.StockCalcWAdays,T.ScheduleCalcMonths, " & _
            " T.DAYSFORSAFETYSTOCK, T.SUMOFRELEASEQTY From Schedule_Upload_PKC_hdr H with (nolock) inner join Schedule_Upload_PKC_Dtl D with (nolock)" & _
            " on h.unit_code = d.unit_code and h.doc_no = d.doc_no inner join CUSTITEM_MST C with (nolock)" & _
            " on c.unit_code = h.unit_code and c.account_code = h.customer_code" & _
             " and c.cust_drgno = d.BuyersPartNumber inner join ITEM_MST I with (nolock) on i.unit_code = c.unit_code and i.item_code = c.item_code" & _
             " inner join TMPSCHEDULESAFETYSTOCK T with (nolock) on t.unit_code = h.unit_code and t.custdrg_no = c.cust_drgno " & _
             " inner join SCHEDULEPARAMETER_DTL SP with (nolock) on C.Account_code = SP.Consignee_code And C.UNIT_CODE = SP.UNIT_CODE" & _
            " AND SP.CUST_DRGNO = T.CUSTDRG_NO AND SP.UNIT_CODE = T.UNIT_CODE    AND SP.WH_CODE = H.suppliercode " & _
             " Where C.Active = 1 AND C.SCHUPLDREQD = 1 And H.Doc_No= " & TRANS_NUMBER & " and SP.CUSTOMER_CODE  = '" & txtCustomerCode.Text & "'" & _
             " AND D.UNIT_CODE = '" & gstrUNITID & "'  group by D.ScheduleDate,C.Cust_Drgno,I.Item_Code, I.DESCRIPTION,H.suppliercode,T.SAFETYSTOCK ," & _
             " SP.SAFETYDAYSMAX, SP.SAFETYDAYS,SP.Consignee_code , SP.SAFETYDAYSMIN, T.StockCalcWAdays, T.ScheduleCalcMonths, T.DAYSFORSAFETYSTOCK, T.SUMOFRELEASEQTY" & _
             " Order By D.ScheduleDate  "

        sqlCMDDisplay.CommandText = sql
        sqlRdrDisplay = sqlCMDDisplay.ExecuteReader

        SCHQTY = 0

        Dim COUNT As Integer
        COUNT = 0
        If sqlRdrDisplay.HasRows Then
            rsbagqty = New SqlCommand
            rsbagqty.Connection = SqlConnectionclass.GetConnection
            Rs = New SqlCommand
            Rs.Connection = SqlConnectionclass.GetConnection

            While sqlRdrDisplay.Read

                rsbagqty.CommandText = "select bag_qty from item_mst where item_code = '" & sqlRdrDisplay("Item_Code").ToString & "' and Status = 'A' AND UNIT_CODE = '" & gstrUNITID & "'"
                rdrbagQty = rsbagqty.ExecuteReader
                If rdrbagQty.HasRows Then
                    rdrbagQty.Read()
                    mlngBAGQTY = rdrbagQty("BAG_QTY").ToString
                Else
                    mlngBAGQTY = 1
                End If
                rdrbagQty.Close()

                sql = " Select TransitDaysBySea, BufferDays "
                sql = sql & "  From ScheduleParameter_mst"
                sql = sql & " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and consignee_code = '" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "' AND UNIT_CODE = '" & gstrUNITID & "'"

                rsbagqty.CommandText = sql
                rdrbagQty = rsbagqty.ExecuteReader

                If rdrbagQty.HasRows Then
                    rdrbagQty.Read()
                    Transit_Days = IIf(IsDBNull(rdrbagQty("TransitDaysBySea").ToString), 0, rdrbagQty("TransitDaysBySea").ToString)
                    Buffer_Days = IIf(IsDBNull(rdrbagQty("BufferDays").ToString), 0, rdrbagQty("BufferDays").ToString)
                End If
                rdrbagQty.Close()

                rsDate = New SqlCommand
                rsDate.Connection = SqlConnectionclass.GetConnection

                sql = "set dateformat 'dmy' select max(dt) as dt from Calendar_Mfg_mst" & _
                    " where work_flg=0 AND UNIT_CODE = '" & gstrUNITID & "' and "

                sql = sql + " dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), VB6.Format(CDate(sqlRdrDisplay("ScheduleDate").ToString), "dd MMM yyyy")) & "' "

                rsDate.CommandText = sql
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If

                rdr = rsDate.ExecuteReader
                If rdr.HasRows Then
                    rdr.Read()
                    dtDate = CDate(IIf(rdr("dt").ToString = "", "01 jan 1900", rdr("dt").ToString))
                End If
                rdr.Close()

                If CDate(getDateForDB(dtDate)) >= CDate(getDateForDB(GetServerDate())) Then
                    sql = "SELECT WHSTOCK,MINSTOCK,RCVDQTY,ISSUEDQTY,WHDATE,REVNO FROM TMPWHSTOCK " & _
                         " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                         " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "

                    sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("SupplierCode").ToString & "' "
                    varWHCODE = sqlRdrDisplay("SupplierCode").ToString

                    RSWHSTOCK.GetResult(sql, ADODB.CursorTypeEnum.adOpenStatic)

                    If RSWHSTOCK.GetNoRows > 0 Then
                        varWhStock = RSWHSTOCK.GetValue("WHSTOCK")
                        varIssuedQty = RSWHSTOCK.GetValue("ISSUEDQTY")
                        varRcvdQty = RSWHSTOCK.GetValue("RCVDQTY")
                        varRevNo = RSWHSTOCK.GetValue("REVNO")
                    Else
                        varWhStock = 0
                        varIssuedQty = 0
                        varRcvdQty = 0
                        varRevNo = 0
                    End If

                    If varWhStock < 0 Then
                        sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                                " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                                " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                                " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
                                " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                                " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
                                " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("SupplierCode").ToString & "', " & _
                                " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "',CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                " 0,CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                " '" & -(varWhStock) & "','" & txtCustomerCode.Text & "','" & varWhStock & "'," & _
                                " 0,'" & -(varWhStock) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                                " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                                " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                                " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                                " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                                " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                                " " & Buffer_Days & ",'" & gstrUNITID & "' )"

                        sqlInsertUpdate.CommandText = sql
                        sqlInsertUpdate.ExecuteNonQuery()

                        sql = "UPDATE TMPWHSTOCK SET WHSTOCK = 0 " & _
                            " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                            " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' " & _
                            " AND WH_CODE = '" & sqlRdrDisplay("SupplierCode").ToString & "' "

                        sqlInsertUpdate.CommandText = sql
                        sqlInsertUpdate.ExecuteNonQuery()
                    End If
                End If
                With Rs
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "SCHEDULEQTY_PKC"
                    .CommandTimeout = 0

                    .Parameters.Clear()
                    .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 8, Trim(txtCustomerCode.Text)).Direction = ParameterDirection.Input
                    .Parameters.Add("@CUST_CODE", SqlDbType.VarChar, 8, Trim(txtCustomerCode.Text)).Direction = ParameterDirection.Input
                    .Parameters.Add("@DOCNO", SqlDbType.Int, 12, Trim(TRANS_NUMBER)).Direction = ParameterDirection.Input
                    .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 50, Trim(sqlRdrDisplay("Cust_DrgNo").ToString)).Direction = ParameterDirection.Input

                    .Parameters.Add("@WH_CODE", SqlDbType.VarChar, 12, sqlRdrDisplay("SUPPLIERCODE").ToString).Direction = ParameterDirection.Input
                    .Parameters.Add("@TRANS_DT", SqlDbType.Date, 11, Format(CDate(Trim(sqlRdrDisplay("ScheduleDate"))), "dd MMM yyyy")).Direction = ParameterDirection.Input
                    .Parameters.Add("@SCHQTY", SqlDbType.Int, 12, 0).Direction = ParameterDirection.Output

                    sql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST"
                    sql = sql + " WHERE CUST_VENDOR_CODE = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "'"
                    'sql = sql + " AND DOCK_CODE = '" & sqlRdrDisplay("PARTY_ID3").ToString & "'"
                    sql = sql + " AND UNIT_CODE = '" & gstrUNITID & "'"

                    rsbagqty.CommandText = sql

                    rdrbagQty = rsbagqty.ExecuteReader
                    If rdrbagQty.HasRows Then
                        rdrbagQty.Read()
                        CONSIGNEE_CODE = Trim(rdrbagQty("CUSTOMER_CODE").ToString)
                    Else
                        CONSIGNEE_CODE = ""
                    End If

                    rdrbagQty.Close()

                    .Parameters(0).Value = Trim(gstrUNITID) : .Parameters(1).Value = Trim(txtCustomerCode.Text) : .Parameters(2).Value = Trim(TRANS_NUMBER)
                    .Parameters(3).Value = Trim(sqlRdrDisplay("Cust_DrgNo").ToString) : .Parameters(4).Value = sqlRdrDisplay("SUPPLIERCODE").ToString
                    .Parameters(5).Value = Format(CDate(Trim(sqlRdrDisplay("ScheduleDate"))), "dd MMM yyyy")

                    Rs.ExecuteScalar()

                End With

                sql = "SELECT WHSTOCK,MINSTOCK,RCVDQTY,ISSUEDQTY,WHDATE,REVNO FROM TMPWHSTOCK " & _
                         " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                         " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "

                sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "' "
                varWHCODE = sqlRdrDisplay("SUPPLIERCODE").ToString

                RSWHSTOCK.GetResult(sql, ADODB.CursorTypeEnum.adOpenStatic)

                If RSWHSTOCK.GetNoRows > 0 Then
                    varWhStock = RSWHSTOCK.GetValue("WHSTOCK")
                    varIssuedQty = RSWHSTOCK.GetValue("ISSUEDQTY")
                    varRcvdQty = RSWHSTOCK.GetValue("RCVDQTY")
                    varRevNo = RSWHSTOCK.GetValue("REVNO")
                Else
                    varWhStock = 0
                    varIssuedQty = 0
                    varRcvdQty = 0
                    varRevNo = 0
                End If

                If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(VB6.Format(dtDate, "dd MMM yyyy")), CDate(VB6.Format(GetServerDate(), "dd MMM yyyy"))) > 0 Then
                    SCHQTY = 0
                    GoTo NOSAFETYSTOCKPERDAY
                End If

                If RSWHSTOCK.GetNoRows > 0 Then
                    If Val(varWhStock) + Val(varRcvdQty) > Val(varIssuedQty) Then
                        SCHQTY = varIssuedQty
                    Else
                        If Val(varWhStock) > 0 Then
                            SCHQTY = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)
                        Else
                            SCHQTY = Val(varIssuedQty) - (Val(varWhStock) + Val(varRcvdQty))
                        End If
                    End If
                Else
                    SCHQTY = IIf(IsDBNull(Rs.Parameters("@SCHQTY").Value), 0, Rs.Parameters("@SCHQTY").Value)
                End If

                If chkDlyPullQty.Checked = True Then

                    Dim schqty1 As String

                    schqty1 = FN_DAILYPULLQTY(varWhStock, varIssuedQty, varRcvdQty, SCHQTY, varWHCODE, sqlRdrDisplay("SAFETYDAYS").ToString, CStr(sqlRdrDisplay("Item_Code").ToString), CStr(sqlRdrDisplay("Cust_DrgNo").ToString), CInt(Row), CStr(sqlRdrDisplay("Factory_Code").ToString), mlngBAGQTY)
                    intPos = InStr(1, schqty1, "*")
                    dailypullrate = 0
                    If intPos > 0 Then
                        dailypullrate = Mid(schqty1, intPos + 1, Len(schqty1))
                        SCHQTY = Mid(schqty1, 1, intPos - 1)
                    End If
                    blnDAILYPULLFLAG = 1
                Else
                    'If Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) > 0 Then
                    '    If ((Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) * Val(sqlRdrDisplay("safetydaysmin").ToString)) - SCHQTY) >= 0 Then
                    '        SCHQTY = 0
                    '    Else
                    '        SAFETYDAYS_BELOW = SCHQTY - (Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) * Val(sqlRdrDisplay("safetydaysmin").ToString))  ''/ val(adors!SAFETYSTKPERDAY)
                    '        SFTYDAYS_MNTD = Val(sqlRdrDisplay("safetydaysMAX").ToString) - Val(sqlRdrDisplay("safetydaysmin").ToString)  ''+ SAFETYDAYS_BELOW
                    '        SFTYDAYS_MNTD = System.Math.Round(SFTYDAYS_MNTD, 0)

                    '        SCHQTY = (Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString) * SFTYDAYS_MNTD) + SAFETYDAYS_BELOW

                    '        SCHQTY = System.Math.Round(SCHQTY, 0)

                    '        If Val(mlngBAGQTY) >= Val(SCHQTY) Then
                    '            SCHQTY = mlngBAGQTY
                    '        Else
                    '            If mlngBAGQTY > 0 Then
                    '                If SCHQTY Mod mlngBAGQTY > 0 Then
                    '                    SCHQTY = (mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)) + mlngBAGQTY
                    '                Else
                    '                    SCHQTY = mlngBAGQTY * Int(SCHQTY / mlngBAGQTY)
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                    'Else
                    If SCHQTY > 0 Then
                        If (Val(varRcvdQty) + Val(varWhStock)) >= SCHQTY Then
                            SCHQTY = 0
                            GoTo NOSAFETYSTOCKPERDAY
                        Else
                            SCHQTY = Val(varIssuedQty) - (Val(varRcvdQty) + Val(varWhStock))
                        End If
                    Else
                        SCHQTY = Val(varIssuedQty) - (Val(varRcvdQty) + Val(varWhStock))
                    End If

                    'SFTYDAYS_MNTD = Val(sqlRdrDisplay("SafetyDaysMax").ToString) - Val(sqlRdrDisplay("safetydaysmin").ToString)
                    SFTYDAYS_MNTD = Val(sqlRdrDisplay("safetydaysmin").ToString)
                    SFTYDAYS_MNTD = System.Math.Round(SFTYDAYS_MNTD, 0)

                    SCHQTY = SCHQTY + (SFTYDAYS_MNTD * Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString))

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
                    Else
                        SCHQTY = 0
                    End If

                    'End If
                End If

NOSAFETYSTOCKPERDAY:
                sql = "select top 1 trans_dt,customer_code," & _
                    " revno from warehouse_stock_dtl" & _
                    " where customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                sql = sql & " and  warehouse_code = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "'"
                sql = sql + " group by customer_code, trans_dt,revno" & _
                " order by trans_dt desc,revno desc "

                rsWHDt = New SqlCommand
                rsWHDt.Connection = SqlConnectionclass.GetConnection
                rsWHDt.CommandText = sql
                rdrWHDt = rsWHDt.ExecuteReader

                If rdrWHDt.HasRows Then
                    rdrWHDt.Read()
                    WHDATE = rdrWHDt("TRANS_DT").ToString
                Else
                    WHDATE = ""
                End If

                rdrWHDt.Close()
                rsWHDt.Dispose()
                rsWHDt = Nothing

                If chkDlyPullQty.Checked = True Then
                    blnDAILYPULLFLAG = 1
                Else
                    blnDAILYPULLFLAG = 0
                End If

                If Val(SCHQTY) > 0 Then
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

                If sqlInsertUpdate Is Nothing Then
                    sqlInsertUpdate = New SqlCommand
                    sqlInsertUpdate.Connection = SqlConnectionclass.GetConnection
                End If

                sqlInsertUpdate.CommandText = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                            " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                            " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                            " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
                            " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                            " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
                            " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("SUPPLIERCODE").ToString & "', " & _
                            " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "','" & Format(CDate(sqlRdrDisplay("ScheduleDate").ToString), "dd MMM yyyy") & "'," & _
                            " '" & Val(sqlRdrDisplay("shipqty").ToString) & "','" & Format(dtDate, "dd MMM yyyy") & "'," & _
                            " '" & SCHQTY & "','" & txtCustomerCode.Text & "','" & IIf(IsDBNull(varWhStock), 0, varWhStock) & "'," & _
                            " '" & varRcvdQty & "','" & varIssuedQty & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                            " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                            " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                            " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                            " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                            " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                            " " & Buffer_Days & ",'" & gstrUNITID & "' )"
                sqlInsertUpdate.ExecuteNonQuery()

                TMPWHSTOCK = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)

                If Val(sqlRdrDisplay("shipqty").ToString) >= 0 Then
                    TMPWHSTOCK = TMPWHSTOCK + SCHQTY                        ''- val(adors!SHIPQTY)
                End If
                updSQL = "UPDATE TMPWHSTOCK" & _
                    " SET WHSTOCK = " & TMPWHSTOCK & " , "

                updSQL = updSQL + " WHDATE = '" & VB6.Format(sqlRdrDisplay("ScheduleDate"), "dd MMM yyyy") & "'"

                updSQL = updSQL + " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                    " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "


                updSQL = updSQL + " AND WH_CODE = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "' " '& _
                '    " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("PARTY_ID3").ToString) & "'"


                mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(updSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

SKIP:

            End While

            sqlRdrDisplay.Close()

            If Not sqlInsertUpdate Is Nothing Then
                sqlInsertUpdate.Dispose()
                sqlInsertUpdate = Nothing
            End If
        Else
            sqlRdrDisplay.Close()
        End If

        If Not sqlCMDDisplay Is Nothing Then
            sqlCMDDisplay.Dispose()
            sqlCMDDisplay = Nothing
        End If

        If Not Rs Is Nothing Then
            Rs.Dispose()
            Rs = Nothing
        End If

        If Not rsDate Is Nothing Then
            rsDate.Dispose()
            rsDate = Nothing
        End If

        If Not sqltrans Is Nothing Then
            sqltrans.Commit()
            sqltrans = Nothing
        End If

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Function

ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        If Not sqlCMDDisplay Is Nothing Then
            sqlCMDDisplay.Dispose()
            sqlCMDDisplay = Nothing
        End If

        If Not Rs Is Nothing Then
            Rs.Dispose()
            Rs = Nothing
        End If

        RSWHSTOCK.ResultSetClose() : RSWHSTOCK = Nothing

        If Not rsDate Is Nothing Then
            rsDate.Dispose()
            rsDate = Nothing
        End If

        If Not sqltrans Is Nothing Then
            sqltrans.Rollback()
            sqltrans = Nothing
        End If

        Return ""
        Exit Function

    End Function

    Private Function FN_Display_BOSCH(ByVal TRANS_NUMBER As String, ByVal FileType As String) As String
        On Error GoTo ERR_Renamed

        Dim Transit_Days, Row, Buffer_Days As Integer
        Dim SCHQTY As Long
        Dim sqlCMDDisplay As SqlCommand
        Dim sqlRdrDisplay As SqlDataReader
        Dim SFTYDAYS_MNTD As Object
        Dim SAFETYDAYS_BELOW As Long
        Dim TMPWHSTOCK As Long
        Dim rsDate As SqlCommand
        Dim sql As String = String.Empty, updSQL As String = String.Empty, strWHCode As String = String.Empty
        Dim RSWHSTOCK As New ClsResultSetDB
        Dim rsWHDt As SqlCommand
        Dim rdrWHDt As SqlDataReader
        Dim varWhStock As Object = Nothing
        Dim varIssuedQty As Object = Nothing
        Dim varRcvdQty As Object = Nothing
        Dim varRevNo As Object = Nothing
        Dim rsbagqty As SqlCommand
        Dim rdrbagQty As SqlDataReader
        Dim dailypullrate As Long
        Dim intPos As Integer
        Dim sqlInsertUpdate As SqlCommand
        Dim CONSIGNEE_CODE As String
        Dim varWHCODE As Object = Nothing
        Dim blnDAILYPULLFLAG As Integer
        Dim Rs As SqlCommand
        Dim rdr As SqlDataReader
        Dim sqltrans As SqlTransaction
        Dim rsTransitDays As ADODB.Recordset
        Dim WHDATE As Date
        Dim dtDate As Date

        blnDAILYPULLFLAG = 0

        sqlCMDDisplay = New SqlCommand
        sqlCMDDisplay.Connection = SqlConnectionclass.GetConnection

        sqltrans = sqlCMDDisplay.Connection.BeginTransaction
        sqlCMDDisplay.Transaction = sqltrans

        sql = " Select Distinct D.SCHEDULEDATE,C.CUST_DRGNO,I.ITEM_CODE, I.DESCRIPTION,H.SUPPLIERCODE,T.SAFETYSTOCK AS SAFETYSTKPERDAY , SP.SAFETYDAYSMAX, " & _
                      " SP.SAFETYDAYS,SP.CONSIGNEE_CODE,SP.SAFETYDAYSMIN,sum(DispatchQty) AS SHIPQTY, T.StockCalcWAdays,T.ScheduleCalcMonths," & _
                      " T.DAYSFORSAFETYSTOCK,T.SUMOFRELEASEQTY   From Schedule_Upload_Bosch_Hdr H WITH (NOLOCK)  INNER JOIN Schedule_Upload_Bosch_Dtl D WITH (NOLOCK)  " & _
                      " ON H.UNIT_CODE = D.UNIT_CODE AND H.DOC_NO = D.DOC_NO INNER JOIN CUSTITEM_MST C WITH (NOLOCK) " & _
                      " ON C.UNIT_CODE = H.UNIT_CODE AND C.ACCOUNT_CODE = H.CUST_CODE AND C.CUST_DRGNO = D.BUYERSPARTNUMBER INNER JOIN ITEM_MST I WITH (NOLOCK) " & _
                      " ON I.UNIT_CODE = C.UNIT_CODE AND I.ITEM_CODE = C.ITEM_CODE INNER JOIN TMPSCHEDULESAFETYSTOCK T WITH (NOLOCK)" & _
                      " ON T.UNIT_CODE = H.UNIT_CODE AND T.CUSTDRG_NO = C.CUST_DRGNO  INNER JOIN SCHEDULEPARAMETER_DTL SP WITH (NOLOCK) " & _
                      " ON C.ACCOUNT_CODE = SP.CONSIGNEE_CODE AND C.UNIT_CODE = SP.UNIT_CODE AND SP.CUST_DRGNO = T.CUSTDRG_NO AND SP.UNIT_CODE = T.UNIT_CODE    AND SP.WH_CODE = H.SUPPLIERCODE" & _
                      " WHERE C.ACTIVE = 1 AND C.SCHUPLDREQD = 1  AND H.DOC_NO=" & TRANS_NUMBER & "  AND SP.CUSTOMER_CODE  =  '" & txtCustomerCode.Text & "'" & _
                      " AND D.UNIT_CODE ='" & gstrUNITID & "'  GROUP BY D.SCHEDULEDATE,C.CUST_DRGNO,I.ITEM_CODE, I.DESCRIPTION,H.SUPPLIERCODE,T.SAFETYSTOCK ,SP.SAFETYDAYSMAX," & _
                      " SP.SAFETYDAYS,SP.CONSIGNEE_CODE , SP.SAFETYDAYSMIN, T.STOCKCALCWADAYS, T.SCHEDULECALCMONTHS, T.DAYSFORSAFETYSTOCK, T.SUMOFRELEASEQTY  " & _
                      " ORDER BY D.SCHEDULEDATE"

        sqlCMDDisplay.CommandText = sql
        sqlRdrDisplay = sqlCMDDisplay.ExecuteReader

        SCHQTY = 0

        Dim COUNT As Integer
        COUNT = 0
        If sqlRdrDisplay.HasRows Then
            rsbagqty = New SqlCommand
            rsbagqty.Connection = SqlConnectionclass.GetConnection
            Rs = New SqlCommand
            Rs.Connection = SqlConnectionclass.GetConnection

            While sqlRdrDisplay.Read

                rsbagqty.CommandText = "select bag_qty from item_mst where item_code = '" & sqlRdrDisplay("Item_Code").ToString & "' and Status = 'A' AND UNIT_CODE = '" & gstrUNITID & "'"
                rdrbagQty = rsbagqty.ExecuteReader
                If rdrbagQty.HasRows Then
                    rdrbagQty.Read()
                    mlngBAGQTY = rdrbagQty("BAG_QTY").ToString
                Else
                    mlngBAGQTY = 1
                End If
                rdrbagQty.Close()

                sql = " Select TransitDaysBySea, BufferDays "
                sql = sql & "  From ScheduleParameter_mst"
                sql = sql & " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and consignee_code = '" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "' AND UNIT_CODE = '" & gstrUNITID & "'"

                rsbagqty.CommandText = sql
                rdrbagQty = rsbagqty.ExecuteReader

                If rdrbagQty.HasRows Then
                    rdrbagQty.Read()
                    Transit_Days = IIf(IsDBNull(rdrbagQty("TransitDaysBySea").ToString), 0, rdrbagQty("TransitDaysBySea").ToString)
                    Buffer_Days = IIf(IsDBNull(rdrbagQty("BufferDays").ToString), 0, rdrbagQty("BufferDays").ToString)
                End If
                rdrbagQty.Close()

                rsDate = New SqlCommand
                rsDate.Connection = SqlConnectionclass.GetConnection

                sql = "set dateformat 'dmy' select max(dt) as dt from Calendar_Mfg_mst" & _
                    " where work_flg=0 AND UNIT_CODE = '" & gstrUNITID & "' and "

                sql = sql + " dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), VB6.Format(CDate(sqlRdrDisplay("ScheduleDate").ToString), "dd MMM yyyy")) & "' "

                rsDate.CommandText = sql
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If

                rdr = rsDate.ExecuteReader
                If rdr.HasRows Then
                    rdr.Read()
                    dtDate = CDate(IIf(rdr("dt").ToString = "", "01 jan 1900", rdr("dt").ToString))
                End If
                rdr.Close()

                If CDate(getDateForDB(dtDate)) >= CDate(getDateForDB(GetServerDate())) Then
                    sql = "SELECT WHSTOCK,MINSTOCK,RCVDQTY,ISSUEDQTY,WHDATE,REVNO FROM TMPWHSTOCK " & _
                         " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                         " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "

                    sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("SupplierCode").ToString & "' "
                    varWHCODE = sqlRdrDisplay("SupplierCode").ToString

                    RSWHSTOCK.GetResult(sql, ADODB.CursorTypeEnum.adOpenStatic)

                    If RSWHSTOCK.GetNoRows > 0 Then
                        varWhStock = RSWHSTOCK.GetValue("WHSTOCK")
                        varIssuedQty = RSWHSTOCK.GetValue("ISSUEDQTY")
                        varRcvdQty = RSWHSTOCK.GetValue("RCVDQTY")
                        varRevNo = RSWHSTOCK.GetValue("REVNO")
                    Else
                        varWhStock = 0
                        varIssuedQty = 0
                        varRcvdQty = 0
                        varRevNo = 0
                    End If

                    If varWhStock < 0 Then
                        sql = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                                " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                                " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                                " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
                                " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                                " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
                                " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("SupplierCode").ToString & "', " & _
                                " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "',CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                " 0,CONVERT(DATETIME,convert(varchar(11),getdate(),106),106)," & _
                                " '" & -(varWhStock) & "','" & txtCustomerCode.Text & "','" & varWhStock & "'," & _
                                " 0,'" & -(varWhStock) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                                " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                                " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                                " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                                " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                                " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                                " " & Buffer_Days & ",'" & gstrUNITID & "' )"

                        sqlInsertUpdate.CommandText = sql
                        sqlInsertUpdate.ExecuteNonQuery()

                        sql = "UPDATE TMPWHSTOCK SET WHSTOCK = 0 " & _
                            " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                            " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' " & _
                            " AND WH_CODE = '" & sqlRdrDisplay("SupplierCode").ToString & "' "

                        sqlInsertUpdate.CommandText = sql
                        sqlInsertUpdate.ExecuteNonQuery()
                    End If
                End If
                With Rs
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "SCHEDULEQTY_BOSCH"
                    .CommandTimeout = 0

                    .Parameters.Clear()
                    .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 8, Trim(txtCustomerCode.Text)).Direction = ParameterDirection.Input
                    .Parameters.Add("@CUST_CODE", SqlDbType.VarChar, 8, Trim(txtCustomerCode.Text)).Direction = ParameterDirection.Input
                    .Parameters.Add("@DOCNO", SqlDbType.Int, 12, Trim(TRANS_NUMBER)).Direction = ParameterDirection.Input
                    .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 50, Trim(sqlRdrDisplay("Cust_DrgNo").ToString)).Direction = ParameterDirection.Input

                    .Parameters.Add("@WH_CODE", SqlDbType.VarChar, 12, sqlRdrDisplay("SUPPLIERCODE").ToString).Direction = ParameterDirection.Input
                    .Parameters.Add("@TRANS_DT", SqlDbType.Date, 11, Format(CDate(Trim(sqlRdrDisplay("ScheduleDate"))), "dd MMM yyyy")).Direction = ParameterDirection.Input
                    .Parameters.Add("@SCHQTY", SqlDbType.Int, 12, 0).Direction = ParameterDirection.Output

                    sql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST"
                    sql = sql + " WHERE CUST_VENDOR_CODE = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "'"
                    'sql = sql + " AND DOCK_CODE = '" & sqlRdrDisplay("PARTY_ID3").ToString & "'"
                    sql = sql + " AND UNIT_CODE = '" & gstrUNITID & "'"

                    rsbagqty.CommandText = sql

                    rdrbagQty = rsbagqty.ExecuteReader
                    If rdrbagQty.HasRows Then
                        rdrbagQty.Read()
                        CONSIGNEE_CODE = Trim(rdrbagQty("CUSTOMER_CODE").ToString)
                    Else
                        CONSIGNEE_CODE = ""
                    End If

                    rdrbagQty.Close()

                    .Parameters(0).Value = Trim(gstrUNITID) : .Parameters(1).Value = Trim(txtCustomerCode.Text) : .Parameters(2).Value = Trim(TRANS_NUMBER)
                    .Parameters(3).Value = Trim(sqlRdrDisplay("Cust_DrgNo").ToString) : .Parameters(4).Value = sqlRdrDisplay("SUPPLIERCODE").ToString
                    .Parameters(5).Value = Format(CDate(Trim(sqlRdrDisplay("ScheduleDate"))), "dd MMM yyyy")

                    Rs.ExecuteScalar()

                End With

                sql = "SELECT WHSTOCK,MINSTOCK,RCVDQTY,ISSUEDQTY,WHDATE,REVNO FROM TMPWHSTOCK " & _
                         " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                         " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "

                sql = sql + " AND WH_CODE = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "' "
                varWHCODE = sqlRdrDisplay("SUPPLIERCODE").ToString

                RSWHSTOCK.GetResult(sql, ADODB.CursorTypeEnum.adOpenStatic)

                If RSWHSTOCK.GetNoRows > 0 Then
                    varWhStock = RSWHSTOCK.GetValue("WHSTOCK")
                    varIssuedQty = RSWHSTOCK.GetValue("ISSUEDQTY")
                    varRcvdQty = RSWHSTOCK.GetValue("RCVDQTY")
                    varRevNo = RSWHSTOCK.GetValue("REVNO")
                Else
                    varWhStock = 0
                    varIssuedQty = 0
                    varRcvdQty = 0
                    varRevNo = 0
                End If

                If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(VB6.Format(dtDate, "dd MMM yyyy")), CDate(VB6.Format(GetServerDate(), "dd MMM yyyy"))) > 0 Then
                    SCHQTY = 0
                    GoTo NOSAFETYSTOCKPERDAY
                End If

                If RSWHSTOCK.GetNoRows > 0 Then
                    If Val(varWhStock) + Val(varRcvdQty) > Val(varIssuedQty) Then
                        SCHQTY = varIssuedQty
                    Else
                        If Val(varWhStock) > 0 Then
                            SCHQTY = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)
                        Else
                            SCHQTY = Val(varIssuedQty) - (Val(varWhStock) + Val(varRcvdQty))
                        End If
                    End If
                Else
                    SCHQTY = IIf(IsDBNull(Rs.Parameters("@SCHQTY").Value), 0, Rs.Parameters("@SCHQTY").Value)
                End If

                If chkDlyPullQty.Checked = True Then

                    Dim schqty1 As String

                    schqty1 = FN_DAILYPULLQTY(varWhStock, varIssuedQty, varRcvdQty, SCHQTY, varWHCODE, sqlRdrDisplay("SAFETYDAYS").ToString, CStr(sqlRdrDisplay("Item_Code").ToString), CStr(sqlRdrDisplay("Cust_DrgNo").ToString), CInt(Row), CStr(sqlRdrDisplay("Factory_Code").ToString), mlngBAGQTY)
                    intPos = InStr(1, schqty1, "*")
                    dailypullrate = 0
                    If intPos > 0 Then
                        dailypullrate = Mid(schqty1, intPos + 1, Len(schqty1))
                        SCHQTY = Mid(schqty1, 1, intPos - 1)
                    End If
                    blnDAILYPULLFLAG = 1
                Else
                    If SCHQTY > 0 Then
                        If (Val(varRcvdQty) + Val(varWhStock)) >= SCHQTY Then
                            SCHQTY = 0
                            GoTo NOSAFETYSTOCKPERDAY
                        Else
                            SCHQTY = Val(varIssuedQty) - (Val(varRcvdQty) + Val(varWhStock))
                        End If
                    Else
                        SCHQTY = Val(varIssuedQty) - (Val(varRcvdQty) + Val(varWhStock))
                    End If

                    SFTYDAYS_MNTD = Val(sqlRdrDisplay("safetydaysmin").ToString)
                    SFTYDAYS_MNTD = System.Math.Round(SFTYDAYS_MNTD, 0)

                    SCHQTY = SCHQTY + (SFTYDAYS_MNTD * Val(sqlRdrDisplay("SAFETYSTKPERDAY").ToString))

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
                    Else
                        SCHQTY = 0
                    End If

                End If

NOSAFETYSTOCKPERDAY:
                sql = "select top 1 trans_dt,customer_code," & _
                    " revno from warehouse_stock_dtl" & _
                    " where customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                sql = sql & " and  warehouse_code = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "'"
                sql = sql + " group by customer_code, trans_dt,revno" & _
                " order by trans_dt desc,revno desc "

                rsWHDt = New SqlCommand
                rsWHDt.Connection = SqlConnectionclass.GetConnection
                rsWHDt.CommandText = sql
                rdrWHDt = rsWHDt.ExecuteReader

                If rdrWHDt.HasRows Then
                    rdrWHDt.Read()
                    WHDATE = rdrWHDt("TRANS_DT").ToString
                Else
                    WHDATE = ""
                End If

                rdrWHDt.Close()
                rsWHDt.Dispose()
                rsWHDt = Nothing

                If chkDlyPullQty.Checked = True Then
                    blnDAILYPULLFLAG = 1
                Else
                    blnDAILYPULLFLAG = 0
                End If

                If Val(SCHQTY) > 0 Then
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

                If sqlInsertUpdate Is Nothing Then
                    sqlInsertUpdate = New SqlCommand
                    sqlInsertUpdate.Connection = SqlConnectionclass.GetConnection
                End If

                sqlInsertUpdate.CommandText = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                            " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                            " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                            " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
                            " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
                            " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
                            " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("SUPPLIERCODE").ToString & "', " & _
                            " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "','" & Format(CDate(sqlRdrDisplay("ScheduleDate").ToString), "dd MMM yyyy") & "'," & _
                            " '" & Val(sqlRdrDisplay("shipqty").ToString) & "','" & Format(dtDate, "dd MMM yyyy") & "'," & _
                            " '" & SCHQTY & "','" & txtCustomerCode.Text & "','" & IIf(IsDBNull(varWhStock), 0, varWhStock) & "'," & _
                            " '" & varRcvdQty & "','" & varIssuedQty & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                            " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
                            " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
                            " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
                            " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
                            " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
                            " " & Buffer_Days & ",'" & gstrUNITID & "' )"
                sqlInsertUpdate.ExecuteNonQuery()

                TMPWHSTOCK = Val(varWhStock) + Val(varRcvdQty) - Val(varIssuedQty)

                If Val(sqlRdrDisplay("shipqty").ToString) >= 0 Then
                    TMPWHSTOCK = TMPWHSTOCK + SCHQTY                        ''- val(adors!SHIPQTY)
                End If
                updSQL = "UPDATE TMPWHSTOCK" & _
                    " SET WHSTOCK = " & TMPWHSTOCK & " , "

                updSQL = updSQL + " WHDATE = '" & VB6.Format(sqlRdrDisplay("ScheduleDate"), "dd MMM yyyy") & "'"

                updSQL = updSQL + " WHERE ITEMCODE = '" & sqlRdrDisplay("Cust_DrgNo").ToString & "' " & _
                    " AND CUST_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "


                updSQL = updSQL + " AND WH_CODE = '" & sqlRdrDisplay("SUPPLIERCODE").ToString & "' " '& _
                '    " AND FACTORY_CODE = '" & Trim(sqlRdrDisplay("PARTY_ID3").ToString) & "'"


                mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(updSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

SKIP:

            End While

            sqlRdrDisplay.Close()

            If Not sqlInsertUpdate Is Nothing Then
                sqlInsertUpdate.Dispose()
                sqlInsertUpdate = Nothing
            End If
        Else
            sqlRdrDisplay.Close()
        End If

        If Not sqlCMDDisplay Is Nothing Then
            sqlCMDDisplay.Dispose()
            sqlCMDDisplay = Nothing
        End If

        If Not Rs Is Nothing Then
            Rs.Dispose()
            Rs = Nothing
        End If

        If Not rsDate Is Nothing Then
            rsDate.Dispose()
            rsDate = Nothing
        End If

        If Not sqltrans Is Nothing Then
            sqltrans.Commit()
            sqltrans = Nothing
        End If

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Function

ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        If Not sqlCMDDisplay Is Nothing Then
            sqlCMDDisplay.Dispose()
            sqlCMDDisplay = Nothing
        End If

        If Not Rs Is Nothing Then
            Rs.Dispose()
            Rs = Nothing
        End If

        RSWHSTOCK.ResultSetClose() : RSWHSTOCK = Nothing

        If Not rsDate Is Nothing Then
            rsDate.Dispose()
            rsDate = Nothing
        End If

        If Not sqltrans Is Nothing Then
            sqltrans.Rollback()
            sqltrans = Nothing
        End If

        Return ""
        Exit Function

    End Function


    Public Function FN_Find_Revision() As Short

        On Error GoTo ERR_Renamed

        Dim sql As String
        Dim adors As New ADODB.Recordset
        If Me.OptReleaseFile.Checked = True Then
            sql = " Select max(RevNo) + 1 as RevNO"
            sql = sql & " From Schedule_Upload_Darwin_Hdr"
            sql = sql & " Where Cust_Code ='" & Trim(txtCustomerCode.Text) & "' and consignee_code = '" & txtConsignee.Text & "'"
            sql = sql & " And Ent_dt=getdate() AND UNIT_CODE = '" & gstrUNITID & "'"
        Else
            sql = " Select max(RevNo) + 1 as RevNO"
            sql = sql & " From WareHouse_Stock_Dtl with (nolock)"
            sql = sql & " Where Customer_Code ='" & Trim(txtCustomerCode.Text) & "'"
            sql = sql & " and Warehouse_Code = '" & Trim(Me.txtUnitCode.Text) & "' and consignee_code='" & Trim(txtConsignee.Text) & "'"
            sql = sql & " and trans_dt = '" & VB6.Format(Trim(Me.DTPicker1.Value), "DD MMM YYYY") & "' AND UNIT_CODE = '" & gstrUNITID & "'"

        End If

        adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        If adors.EOF = False Then
            FN_Find_Revision = IIf(IsDBNull(adors.Fields("RevNo").Value), 1, adors.Fields("RevNo").Value)
        Else
            FN_Find_Revision = 1
        End If
        adors.Close()
        adors = Nothing

        Exit Function
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Function

    End Function

    Private Sub FN_WareHouse_File_Upload()

        On Error GoTo ERR_Renamed

        Dim Item_Prefix, Item_Suffix As String
        Dim sql As String
        Dim Col, Row As Short
        Dim Trans_Satus As Boolean
        Dim Rev_No As Short
        Dim lngStockQty As Integer
        Dim sqlCmd As SqlCommand
        Dim sqlRDR As SqlDataReader
        Dim Msg As String
        Dim Flag As Short

        Dim Item_Rate As Double

        Dim WhStkObj As New prj_uploadInvoiceDaimler.prj_uploadInvoiceDaimler          'eMpro-20090309-28458

        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection

        sql = "Delete from warehouse_stock_dtl_temp where UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlCmd.ExecuteNonQuery()

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

        If ChkTextFile.Checked = True And UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) <> UCase("txt") Then
            MsgBox("File Is Not In txt Format.", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        If UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) <> UCase("txt") Then
            Obj_EX = New Excel.Application
            Obj_EX.Workbooks.Open(Trim(Me.txtFileName.Text))
        End If

        If OptStock.Checked = True And UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) = UCase("txt") Then
            Rev_No = FN_Find_Revision()

            Msg = ""

            If ChkFord.Enabled = True And ChkFord.Checked = True Then
                Msg = WhStkObj.FN_WH_TextFileUpload_FORD(txtCustomerCode.Text, txtUnitCode.Text, txtConsignee.Text, DTPicker1.Value, gstrConnectSQLClient, txtFileName.Text, Rev_No, mP_User)
            End If

            If ChkDaimler.Enabled = True And ChkDaimler.Checked = True Then
                Msg = WhStkObj.FN_WareHouse_TextFileUpload(txtCustomerCode.Text, txtUnitCode.Text, txtConsignee.Text, DTPicker1.Value, gstrConnectSQLClient, txtFileName.Text, Rev_No, mP_User)
            End If

            MsgBox(Mid(Msg, 3, Msg.Length))

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

            While Len(Item_Suffix) <> 0
                Row = 3
                sqlCmd.CommandText = "set dateformat 'dmy'"
                sqlCmd.ExecuteNonQuery()

                sql = "INSERT INTO WAREHOUSE_STOCK_DTL_temp(CUSTOMER_CODE, " & "WAREHOUSE_CODE,Consignee_Code,UPLOAD_FILE_NAME,ITEM_CODE,QTY,RATE, " & "TRANS_DT,ENT_DT,ENT_ID, REVNO,UNIT_CODE )" & "VALUES ('" & Me.txtCustomerCode.Text & "','" & Me.txtUnitCode.Text & "'  ,'" & Me.txtConsignee.Text & "'," & " '" & Me.txtFileName.Text & "','" & Item_Prefix & "" & Item_Suffix & "' ," & " " & lngStockQty & "," & Item_Rate & ",'" & VB6.Format(DTPicker1.Value, "dd/MMM/yyyy") & "', " & "  getDate() , '" & mP_User & "'," & Rev_No & ",'" & gstrUNITID & "' )"
                sqlCmd.CommandText = sql
                sqlCmd.ExecuteNonQuery()

                If Col = Obj_EX.Columns.Count Then
                    Exit While
                Else
                    Col = Col + 1
                End If

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

            sql = "DELETE FROM WAREHOUSE_STOCK_DTL_temp " & _
                 "WHERE ITEM_CODE NOT IN (SELECT CUST_DRGNO FROM CUSTITEM_MST " & _
                 "WHERE ACCOUNT_CODE = '" & txtConsignee.Text & "' " & _
                 "AND SCHUPLDREQD = 1 AND ACTIVE=1 AND UNIT_CODE = '" & gstrUNITID & "' ) AND CUSTomer_CODE = '" & txtCustomerCode.Text & "' " & _
                 "and consignee_code = '" & txtConsignee.Text & "' and revno = " & Rev_No & " AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "select  ltrim(rtrim(item_code)) + '     ' + ltrim(rtrim(cust_drgno))" & _
             " as item_custdrgno" & " From custitem_mst " & _
             " where account_code = '" & txtConsignee.Text & "' " & " " & _
             " and cust_drgno in (select distinct item_code from warehouse_stock_dtl_temp" & _
             " with (nolock)" & " where revno = " & Rev_No & " and" & _
             " customer_code = '" & txtCustomerCode.Text & "' and" & "" & _
             " consignee_code = '" & txtConsignee.Text & "' " & " and" & _
             " trans_dt = '" & VB6.Format(Me.DTPicker1.Value, "dd MMM yyyy") & "' AND UNIT_CODE = '" & gstrUNITID & "')" & " and" & _
             " ltrim(rtrim(item_code)) + '     ' + ltrim(rtrim(cust_drgno))" & _
             " NOT in " & " (select ltrim(rtrim(item_code)) + '     ' + ltrim(rtrim(cust_drgno))" & _
             " as item_custdrgno " & " From custitem_mst where" & _
             " account_code = '" & txtCustomerCode.Text & "' AND ACTIVE = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "') AND ACTIVE = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "'"

            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader
            Msg = ""
            If sqlRDR.HasRows Then
                While sqlRDR.Read
                    Msg = Msg & "  " & vbCrLf + sqlRDR("item_custdrgno").ToString
                End While

                MsgBox("Following Items Are Not Defined In The Customer Item Master : " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                Flag = 1
            End If
            sqlRDR.Close()

            sql = " Select distinct W.item_code  from WAREHOUSE_STOCK_dtl_temp W"
            sql = sql & " Where LTrim(RTrim(w.Item_Code))"
            sql = sql & " not in (select Cust_DrgNo  from ScheduleParameter_dtl where  Customer_code = '" & Trim(txtCustomerCode.Text) & "'  AND CONSIGNEE_CODE = '" & Trim(txtConsignee.Text) & "'  And WH_Code = '" & Me.txtUnitCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "')"
            sql = sql & " and W.revno = " & Rev_No & " and W.customer_code = '" & Trim(txtCustomerCode.Text) & "' AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' and w.warehouse_code = '" & Me.txtUnitCode.Text & "' and w.trans_dt = '" & VB6.Format(Me.DTPicker1.Value, "dd MMM yyyy") & "' AND W.UNIT_CODE = '" & gstrUNITID & "'"
            Msg = ""

            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader
            If sqlRDR.HasRows Then
                While sqlRDR.Read
                    Msg = Msg & "  " & vbCrLf + sqlRDR("item_code").ToString
                End While

                MsgBox("These Items Are Not Defined In The Schedule Parameter.: " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                Flag = 1
            End If
            sqlRDR.Close()

            sql = "select c.account_code, c.cust_drgno, count(*) countitem " & " from custitem_mst c where Exists(   Select * " & " from warehouse_stock_dtl_temp w Where w.Item_Code = c.Cust_DrgNo" & " and w.trans_dt = '" & VB6.Format(Me.DTPicker1.Value, "dd MMM yyyy") & "' and w.revno = '" & Rev_No & "' " & " and c.account_code = w.CONSIGNEE_code AND W.UNIT_CODE = '" & gstrUNITID & "' )" & " and     c.active = 1 AND SCHUPLDREQD = 1 and c.account_code = '" & Me.txtConsignee.Text & "' AND C.UNIT_CODE = '" & gstrUNITID & "'" & " group by account_code, cust_drgno "
            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader

            Msg = ""

            If sqlRDR.HasRows Then
                While sqlRDR.Read
                    countREC = sqlRDR("COUNTITEM").ToString
                    If countREC > 1 Then
                        Msg = Msg & vbCrLf + sqlRDR("cust_drgno").ToString
                        Flag = 1
                    End If
                End While
                If Msg <> "" Then
                    MsgBox("For Consignee: " & txtConsignee.Text & " Following Cust_DrgNo Are Active For Multiple Items : " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            End If
            sqlRDR.Close()
            If Flag = 1 Then
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            Else
                sql = "SELECT * FROM WAREHOUSE_STOCK_DTL_temp " & _
                    " WHERE WAREHOUSE_CODE = '" & txtUnitCode.Text & "'" & _
                    " AND CUSTomer_CODE = '" & txtCustomerCode.Text & "' " & _
                    " AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' " & _
                    " AND revno = " & Rev_No & " and trans_dt = '" & VB6.Format(DTPicker1.Value, "dd MMM yyyy") & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                sqlCmd.CommandText = sql
                sqlRDR = sqlCmd.ExecuteReader

                If Not sqlRDR.HasRows Then
                    MsgBox("Warehouse Stock Not Uploaded As No Item Defined In The System", MsgBoxStyle.Information, ResolveResString(100))
                    sqlRDR.Close()
                Else
                    sqlRDR.Close()
                    sql = "insert into WAREHOUSE_STOCK_DTL select * from WAREHOUSE_STOCK_DTL_temp WHERE  UNIT_CODE = '" & gstrUNITID & "'"
                    sqlCmd.CommandText = sql
                    sqlCmd.ExecuteNonQuery()
                    MsgBox("WareHouse Stock Uploaded Succesfully !", MsgBoxStyle.Information, ResolveResString(100))

                    Obj_FSO = Nothing

                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                End If
            End If
        End If

        sqlCmd.Dispose()
        sqlCmd = Nothing

        If OptRecvd.Checked = True Then
            Call WareHouse_Inv_Upload()
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
        End If
        Exit Sub

ERR_Renamed:

        If Trans_Satus = True Then
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Exit Sub
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub

    Private Sub chkDlyPullQty_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDlyPullQty.CheckStateChanged
        On Error GoTo ERR_Renamed

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

        Exit Sub
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Private Sub CmdClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdClear.Click

        On Error GoTo ERR_Renamed
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
        ChkTextFile.Checked = False
        ChkDaimler.Checked = False
        ChkFord.Checked = False

        lblMessage1.Text = ""
        lblMessage2.Text = ""
        lblMessage3.Text = ""

        Exit Sub
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        On Error GoTo ErrHandler
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

    End Sub

    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustHelp.Click

        On Error GoTo ErrHandler

        Dim strCustHelp() As String = Nothing
        Dim Rs As New ClsResultSetDB
        Call CmdClear_Click(CmdClear, New System.EventArgs())
        mblnDailymktUpdated = False
        mblnfilemove = False

        If OptWareHouseFile.Checked = True Then
            strCustHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code, c.cust_name from customer_mst c, " & "scheduleparameter_mst S where c.customer_code = s.customer_code AND C.UNIT_CODE = S.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'", "List of Customers")

        ElseIf OptReleaseFile.Checked = True Then
            strCustHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code,c.cust_name from customer_mst c," & " scheduleparameter_mst s where c.customer_code = s.customer_code AND C.UNIT_CODE = S.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'", " List of Customer ")
        End If

        If UBound(strCustHelp) <> -1 Then
            If strCustHelp(0) <> "0" Then
                Me.txtCustomerCode.Text = strCustHelp(0)
                Me.LblCustomerName.Text = strCustHelp(1)
                Rs = New ClsResultSetDB
                If OptReleaseFile.Checked Then
                    Rs.GetResult("select top 1 ReleaseFile_Location" & " from scheduleparameter_mst" & " where customer_code = '" & txtCustomerCode.Text & "' and UNIT_CODE = '" & gstrUNITID & "'" & " order by entdt")
                    txtFileName.Text = Rs.GetValue("ReleaseFile_Location")
                    Rs.ResultSetClose()
                    Rs = New ClsResultSetDB
                    Rs.GetResult("Select plant_c,plant_nm from plant_mst WHERE UNIT_CODE ='" & gstrUNITID & "'")
                    txtUnitCode.Text = Rs.GetValue("plant_c")
                    lblUnitName.Text = Rs.GetValue("plant_nm")
                    Rs.ResultSetClose()
                    Call CmdUploadCSV_Click(CmdUploadCSV, New System.EventArgs())
                    CmdClear.Focus()
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdFileHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFileHelp.Click
        On Error GoTo ErrHandler
        Dim sql As String
        Dim rsPath As New ADODB.Recordset
        CommanDLogOpen.FileName = ""
        CommanDLogOpen.InitialDirectory = ""

        If Me.OptReleaseFile.Checked = True Then

            sql = "SELECT RELEASEFILE_LOCATION FROM SCHEDULEPARAMETER_MST " & "WHERE CUSTOMER_CODE = '" & Me.txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"

            rsPath.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
            If rsPath.EOF = False Then
                CommanDLogOpen.InitialDirectory = rsPath.Fields("ReleaseFile_Location").Value
            Else
                MsgBox("No Location Defined")
                CommanDLogOpen.FileName = ""
                CommanDLogOpen.InitialDirectory = gstrLocalCDrive
            End If
        Else

            sql = "SELECT WAREHOUSEFILE_LOCATION FROM SCHEDULEPARAMETER_MST " & "WHERE CUSTOMER_CODE = '" & Me.txtCustomerCode.Text & "'  " & "AND WH_CODE = '" & Me.txtUnitCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' "
            rsPath.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
            If rsPath.EOF = False Then
                CommanDLogOpen.InitialDirectory = rsPath.Fields("WarehouseFile_Location").Value
            Else
                MsgBox("No Location Defined")
                CommanDLogOpen.FileName = ""
                CommanDLogOpen.InitialDirectory = gstrLocalCDrive
            End If

        End If

        If ChkTextFile.Checked = False Then
            CommanDLogOpen.Filter = "Microsoft Excel File (*.xls)|*.xls;*.xlsx;*.CSV"
        End If
        If ChkTextFile.Checked = True Then
            CommanDLogOpen.Filter = "Text Documents (*.Txt)|*.Txt"
        End If
        CommanDLogOpen.ShowDialog()
        Me.txtFileName.Text = CommanDLogOpen.FileName
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub Updt_DailyMkt(ByRef FileType As String)
        On Error GoTo ErrHandler
        Dim sqlCmd As SqlCommand
        Dim intRETVAL As Integer = 0

        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection
        sqlCmd.CommandType = CommandType.StoredProcedure
        sqlCmd.CommandTimeout = 0
        With sqlCmd
            .CommandText = "updt_dailymkt_cdp"
            .Parameters.Clear()
            .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 10, gstrUNITID).Direction = ParameterDirection.Input
            .Parameters.Add("@CUSTOMERCODE", SqlDbType.VarChar, 10, txtCustomerCode.Text).Direction = ParameterDirection.Input
            .Parameters.Add("@DOCNO", SqlDbType.VarChar, 10, txtDocNo.Text).Direction = ParameterDirection.Input
            .Parameters.Add("@USERID", SqlDbType.VarChar, 10, mP_User).Direction = ParameterDirection.Input
            .Parameters.Add("@FILETYPE", SqlDbType.VarChar, 10, Darwin_FileType).Direction = ParameterDirection.Input
            .Parameters.Add("@RETVAL", SqlDbType.Int, 1, 0).Direction = ParameterDirection.Output

            .Parameters(0).Value = gstrUNITID
            .Parameters(1).Value = txtCustomerCode.Text
            .Parameters(2).Value = txtDocNo.Text
            .Parameters(3).Value = mP_User
            .Parameters(4).Value = Darwin_FileType

            .ExecuteScalar()

        End With
        If intRETVAL = sqlCmd.Parameters("@RETVAL").Value Then
            lblMessage2.Text = "No Schedule Data to Save."
        Else
            lblMessage2.Text = "Schedule Updated Successfully for Planning."
        End If

        sqlCmd.Dispose()
        sqlCmd = Nothing

        Exit Sub
ErrHandler:
        If Not sqlCmd Is Nothing Then
            sqlCmd.Dispose()
            sqlCmd = Nothing
        End If

        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdUnitHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitHelp.Click

        On Error GoTo ErrHandler
        Dim strHelp() As String = Nothing
        Dim rsobject As New ClsResultSetDB

        Call rsobject.GetResult("Select distinct Customer_Mst.Cust_Name,ScheduleParameter_mst.TransitDaysbysea From ScheduleParameter_mst,Customer_Mst  Where Customer_Mst.Customer_Code=ScheduleParameter_mst.Customer_Code And Customer_Mst.Customer_Code = '" & Trim(Me.txtCustomerCode.Text) & "' AND Customer_Mst.UNIT_CODE=ScheduleParameter_mst.UNIT_CODE AND Customer_Mst.UNIT_CODE = '" & gstrUNITID & "'")
        Me.LblCustomerName.Text = rsobject.GetValue("Cust_Name")
        Me.lbltransitdaysvalue.Text = rsobject.GetValue("TransitDaysBySea")

        Me.lblUnitName.Text = CStr(Nothing)
        Me.txtUnitCode.Text = CStr(Nothing)

        If OptReleaseFile.Checked = True Then
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select distinct plant_c,plant_nm from plant_mst WHERE UNIT_CODE = '" & gstrUNITID & "'")
        ElseIf OptWareHouseFile.Checked = True Then
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct C.wh_code,W.WH_DESCRIPTION  from custwarehouse_mst C,WAREHOUSE_MST W " & " where C.customer_code = '" & txtConsignee.Text & "' AND C.WH_CODE = W.WH_CODE " & " and active = 1 AND C.UNIT_CODE = W.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'")
        End If

        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                Me.txtUnitCode.Text = strHelp(0)
                Me.lblUnitName.Text = strHelp(1)
                If OptWareHouseFile.Checked = True Then
                    rsobject.GetResult("select top 1 WarehouseFile_Location" & " from scheduleparameter_mst" & " where customer_code = '" & txtCustomerCode.Text & "' and WH_Code='" & Trim(txtUnitCode.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "'" & " order by entdt")
                    txtFileName.Text = rsobject.GetValue("WarehouseFile_Location")
                    rsobject.ResultSetClose()
                End If
            Else
                MsgBox(" No Warehouse Defined for the selected Consignee.", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Public Sub CmdUploadCSV_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdUploadCSV.Click
        Dim rsobject As New ClsResultSetDB

        On Error GoTo ERR_Renamed
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

            rsobject.GetResult("select C.wh_code,W.WH_DESCRIPTION  from custwarehouse_mst C,WAREHOUSE_MST W " & " where C.customer_code = '" & txtConsignee.Text & "' AND C.WH_CODE = W.WH_CODE and c.wh_code = '" & txtUnitCode.Text & "' " & " and active = 1 AND C.UNIT_CODE = W.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'")

            If rsobject.GetNoRows = 0 Then
                MsgBox("Invalid Warehouse Code", MsgBoxStyle.OkOnly, ResolveResString(100))
                txtUnitCode.Text = ""
                Exit Sub
            End If

            FN_WareHouse_File_Upload()
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

        ElseIf OptReleaseFile.Checked = True Then
            Call FN_FILESELECTION()
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

        End If

        Exit Sub
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Sub

    End Sub

    Private Sub FN_Release_File_Upload_Bosch()
        On Error GoTo ERR_Renamed

        Dim Cell_Data As String = ""
        Dim Row As Object = Nothing
        Dim i As Short = 0
        Dim Data_Row() As String = Nothing
        Dim Trans_Satus As Boolean
        Dim Upload_FileType As String = "", trans_number As String = ""
        Dim sql As String = "", Cell_Data1 As String = ""
        Dim Rev_No As Object = Nothing
        Dim Col As Short = 0
        Dim HOLIDAY As String = ""
        Dim Msg As String
        Dim sql1 As Object = Nothing
        Dim sql2 As String
        Dim msgWH As String = Nothing
        Dim ShipmentFlag As Boolean
        Dim sheetNo As Short
        Dim YesNo As String = Nothing
        Dim WA As String
        Dim sch As Integer
        Dim SftyDays, ItemCode, CustDrgNo, SftyStk, ShpgQty As Object
        Dim ShipDate As Object
        Dim wh_code
        Dim RSWHCOUNT As SqlCommand
        Dim RSWHCOUNTSTOCK As SqlCommand
        Dim RdrWHCOUNT As SqlDataReader
        Dim RdrWHCOUNTSTOCK As SqlDataReader

        Dim sqlCmd As SqlCommand
        Dim sqlRDR As SqlDataReader
        Dim SQLTRANS As SqlTransaction
        Dim ISTRANS As Boolean

        Dim dtFileData As New DataTable

        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlApp As Excel.Application

        Flag = 0
        HOLIDAY = ""
        msgWH = ""

        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection
        sqlCmd.CommandType = CommandType.Text

        sqlCmd.CommandText = "DELETE FROM Schedule_Upload_Bosch_Hdr_temp WHERE UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        sqlCmd.CommandText = "DELETE FROM Schedule_Upload_Bosch_Dtl_temp WHERE UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        Obj_EX = New Excel.Application
        Obj_EX.Workbooks.Open(Trim(Me.txtFileName.Text))

        Row = 1
        range = Obj_EX.Cells(Row, 1)
        If Not range.Value Is Nothing Then
            Cell_Data = Replace((range.Value.ToString), "'", "")
        Else
            Cell_Data = ""
        End If

        If Len(Cell_Data) = 0 Then

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        sql = "Select current_no + 1 as current_no from documenttype_mst where Doc_Type = 9999 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            trans_number = sqlRDR("current_no").ToString
        Else
            MessageBox.Show("Document Number Not Generated.", ResolveResString(100), MessageBoxButtons.OK)
            Exit Sub
        End If
        sqlRDR.Close()

        sql = "SELECT iSnull(SHIPMENTTHRUWH,0) as SHIPMENTTHRUWH FROM CUSTOMER_MST WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            ShipmentFlag = sqlRDR("SHIPMENTTHRUWH").ToString
            sqlRDR.Close()
        Else
            MessageBox.Show("Shipment Flag Not Defined in Customer Master.", ResolveResString(100), MessageBoxButtons.OK)
            sqlRDR.Close()
            Exit Sub
        End If

        Trans_Satus = True

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(txtFileName.Text)
        xlWorkSheet = xlWorkBook.ActiveSheet
        Dim result As Boolean = True
        Dim maxLengthBoschExcelColumn As Integer = [Enum].GetValues(GetType(BoschExcelColumn)).Length
        Dim data As Object(,) = DirectCast(xlWorkSheet.UsedRange.Value2, Object(,))

        For Col1 As Integer = 0 To BoshExcelColumnName.Length - 1
            dtFileData.Columns.Add(BoshExcelColumnName(Col1), GetType(System.String))
        Next

        For row1 As Integer = DataRowIndex To data.GetUpperBound(0)
            Dim newDataRow1 As DataRow = dtFileData.NewRow()
            For col1 As Integer = 1 To maxLengthBoschExcelColumn
                newDataRow1(col1 - 1) = Replace(data(row1, col1), "'", "")
            Next
            dtFileData.Rows.Add(newDataRow1)
        Next

        Msg = ""
        If dtFileData IsNot Nothing AndAlso dtFileData.Rows.Count > 0 Then

            For a As Integer = 0 To dtFileData.Rows.Count - 1
                If a = 0 Then
                    sql = "   INSERT INTO Schedule_Upload_Bosch_Hdr_temp(Doc_no,RcvrEDICode,CallOffNo,CallOffDate,SupplyToBuyerPlantcode,SupplyFromPlantCode,SupplierCode,PortOfDischarge,PortOfDischarge_AddIntDest,ReferenceOrderNumber,PrevDeliveryInstrNo,ScheduleCondition,EntDt,EntBy,UpdDt,UpdBy,Unit_Code,Cust_Code,Consignee_Code,UPLOADFILENAME)"
                    sql = sql + " Values (" + trans_number + ",'" + dtFileData.Rows(a)("RcvrEDICode") + "'," + dtFileData.Rows(a)("CallOffNo") + ",'" + dtFileData.Rows(a)("CallOffDate") + "','" + dtFileData.Rows(a)("SupplyToBuyerPlantcode") + "' ,'" + dtFileData.Rows(a)("SupplyFromPlantCode") + "' ,'" + dtFileData.Rows(a)("SupplierCode") + "' ,'" + dtFileData.Rows(a)("PortOfDischarge") + "','" + dtFileData.Rows(a)("PortOfDischarge_AddIntDest") + "' ,'" + dtFileData.Rows(a)("ReferenceOrderNumber") + "' ,'" + dtFileData.Rows(a)("PrevDeliveryInstrNo") + "' ,'" + dtFileData.Rows(a)("SchCondition") + "',GETDATE(),'" + mP_User + "',GETDATE(),'" + mP_User + "','" + gstrUNITID + "','" + txtCustomerCode.Text + "','" + txtCustomerCode.Text + "','" + txtFileName.Text + "')"
                    sqlCmd.CommandText = sql
                    sqlCmd.ExecuteNonQuery()
                End If

                sql = "   INSERT INTO SCHEDULE_UPLOAD_BOSCH_DTL_TEMP(Doc_no,BuyersPartNumber,  CumQtyReceived,   DispatchQty ,     ScheduleDate,    EntDt,EntBy,UpdDt,UpdBy,Unit_Code,CumQtyStartDate)"
                sql = sql & " Values (" & trans_number & ",'" & dtFileData.Rows(a)("BuyerPartNumber") & "','" & dtFileData.Rows(a)("CumQuantityReceived") & "','" & dtFileData.Rows(a)("DispQty") & "' ,'" & dtFileData.Rows(a)("SchDeliveryDate") & "' ,GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "','" & gstrUNITID & "','" & dtFileData.Rows(a)("CumQtyStartDate") & "')"
                sqlCmd.CommandText = sql
                sqlCmd.ExecuteNonQuery()
            Next
        End If
        'End dataTAble
        
        Dim STRCONS As String
        sql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST WHERE CUST_VENDOR_CODE = '" & dtFileData.Rows(0).Item("SupplierCode") & "' AND DOCK_CODE = '" & dtFileData.Rows(0).Item("SupplierCode") & "' AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            STRCONS = sqlRDR("CUSTOMER_CODE").ToString
        Else
            STRCONS = txtCustomerCode.Text
        End If

        sqlRDR.Close()
        '04 july 2023
        'ONE VALIDATION ADDED : CUSTOMER ITEM MASTER MUST EXIST
        sql = "SELECT DISTINCT BUYERSPARTNUMBER FROM SCHEDULE_UPLOAD_BOSCH_DTL_TEMP ST  where UNIT_CODE = '" & gstrUNITID & "' AND doc_no='" & trans_number & "'" & _
            " AND NOT EXISTS ( " & _
            " SELECT TOP 1 1 FROM CUSTITEM_MST CM WHERE ST.UNIT_CODE = CM.UNIT_CODE AND ST.BUYERSPARTNUMBER=CM.CUST_DRGNO " & _
            " AND UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & STRCONS & "'" & _
            " )"


        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            Msg = Msg & "'" + sqlRDR("BUYERSPARTNUMBER").ToString + "'" + vbCrLf
        End If

        sqlRDR.Close()

        If Len(Trim(Msg)) > 0 Then
            MsgBox("Following Items Are Not Defined In The Customer Item Master " & vbCrLf & "for Customer " & txtCustomerCode.Text & " : " & vbCrLf & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
            Flag = 1
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Exit Sub
        End If

        Trans_Satus = False

        '04 july 2023

        sql = "select cust_drgno FROM CUSTITEM_MST " & _
            " WHERE CUST_DRGNO = '" & dtFileData.Rows(0).Item("BuyerPartNumber") & "'" & " AND active = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "' AND"
        sql = sql & " account_code = '" & Me.txtCustomerCode.Text & "' GROUP BY Cust_Drgno HAVING COUNT(CUST_DRGNO)  > 1"

        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader

        If sqlRDR.HasRows Then
            sqlRDR.Read()
            Msg = Msg & "'" + sqlRDR("CUST_DRGNO").ToString + "'" + vbCrLf
        End If
        sqlRDR.Close()

        'sql = "select distinct d.InternalItemCode " & _
        '        " from Schedule_Upload_Bosch_Dtl_temp d,Schedule_Upload_Bosch_Hdr_temp h " & _
        '        " Where d.doc_no = h.doc_no " & _
        '        " and   D.UNIT_CODE = H.UNIT_CODE and h.doc_no=" & trans_number & " AND d.UNIT_CODE = '" & gstrUNITID & "' and internalitemcode is not NULL" & _
        '        " and ltrim(rtrim(d.InternalItemCode)) " & _
        '        " not in (select cust_drgno from custitem_mst where active = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "' "

        sql = " select distinct d.BuyersPartNumber "
        sql = sql + " from Schedule_Upload_Bosch_Dtl d inner join Schedule_Upload_Bosch_Hdr h on d.doc_no = h.doc_no "
        sql = sql + " and D.UNIT_CODE = H.UNIT_CODE where h.doc_no = " & trans_number & " AND d.UNIT_CODE = '" & gstrUNITID & "'"
        sql = sql + " and not exists (select cust_drgno from custitem_mst c where c.UNIT_CODE = h.Unit_Code and c.Account_Code = '" + txtCustomerCode.Text + "'"
        sql = sql + " and c.Cust_Drgno = d.BuyersPartNumber and c.active = 1 AND c.SCHUPLDREQD = 1)"

        'If ShipmentFlag = True Then
        '    sql = sql + " and account_code in (SELECT DISTINCT CONSIGNEE_CODE FROM SCHEDULE_UPLOAD_DARWINEDIFACT_DTL_temp WHERE DOC_NO = " & trans_number & " AND UNIT_CODE = '" & gstrUNITID & "'))"
        'Else
        '    sql = sql & " and account_code = '" & Me.txtCustomerCode.Text & "')"
        'End If

        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        Msg = ""

        If sqlRDR.HasRows Then
            While sqlRDR.Read
                Msg = Msg & "'" + sqlRDR("InternalItemCode").ToString + "'" + vbCrLf
            End While

            If Len(Trim(Msg)) > 0 Then
                MsgBox("Following Items Are Not Defined In The System" & vbCrLf & "for Customer " & txtCustomerCode.Text & " : " & vbCrLf & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                Flag = 1
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            End If

            Trans_Satus = False

        End If
        sqlRDR.Close()

        HOLIDAY = Replace(HOLIDAY, " ", Chr(13))
        If HOLIDAY <> "" Then
            MsgBox("Following is/are not working day(s):" & vbCrLf & vbCrLf & "Consignee---Date" & vbCrLf & HOLIDAY & vbCrLf)
            Flag = 1
        End If

        If ShipmentFlag = True Then
            sql = "select distinct h.SupplierCode from Schedule_Upload_Bosch_Hdr_temp H," & _
                    " scheduleparameter_mst s where h.SupplierCode  not in(select wh_code " & _
                    " from scheduleparameter_mst s where s.customer_code =  '" & Me.txtCustomerCode.Text & "' AND S.UNIT_CODE = '" & gstrUNITID & "')" & _
                    " and customer_code = '" & Me.txtCustomerCode.Text & "' and doc_no = " & trans_number & " AND S.UNIT_CODE = H.UNIT_CODE AND S.UNIT_CODE = '" & gstrUNITID & "' "
            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader

            If sqlRDR.HasRows Then
                While sqlRDR.Read
                    msgWH = msgWH & "  '" + sqlRDR("SupplierCode").ToString + "'  "
                    Flag = 1
                End While
            End If
            sqlRDR.Close()

            If msgWH <> "" Then
                MsgBox("WRONG WAREHOUSE: " & msgWH, MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
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

        If Flag = 1 Then
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Exit Sub
        Else
            If ShipmentFlag = False And chkDlyPullQty.Checked = True Then
                MsgBox("Shipment For This Customer Is Not Through Warehouse, " & vbCrLf & "So, You Can't Use Daily Pull Qty For Safety Stock Calculations.", vbInformation, ResolveResString(100))
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            Else
                Me.txtDocNo.Text = trans_number
            End If
        End If

        If Flag = 0 Then
            SQLTRANS = sqlCmd.Connection.BeginTransaction()
            sqlCmd.Transaction = SQLTRANS
            ISTRANS = True

            sql = "INSERT INTO Schedule_Upload_Bosch_Hdr SELECT * FROM Schedule_Upload_Bosch_Hdr_temp where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "INSERT INTO Schedule_Upload_Bosch_Dtl (Doc_no,BuyersPartNumber,CumQtyReceived,CumQtyStartDate,DispatchQty,ScheduleDate,EntDt,EntBy,UpdDt,UpdBy,Unit_Code)  SELECT Doc_no,BuyersPartNumber,CumQtyReceived,CumQtyStartDate,DispatchQty,ScheduleDate,EntDt,EntBy,UpdDt,UpdBy,Unit_Code FROM Schedule_Upload_Bosch_Dtl_temp where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "update documenttype_mst set current_no = " & trans_number & " where Doc_Type = 9999 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            SQLTRANS.Commit()
            ISTRANS = False
        End If

        If ShipmentFlag = True Then
            RSWHCOUNT = New SqlCommand
            RSWHCOUNT.Connection = SqlConnectionclass.GetConnection

            RSWHCOUNTSTOCK = New SqlCommand
            RSWHCOUNTSTOCK.Connection = SqlConnectionclass.GetConnection

            sql = "Select count(distinct SupplierCode) COUNT,SupplierCode as WH_CODE from Schedule_Upload_Bosch_Hdr with (nolock) where doc_no='" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "' group by SupplierCode"
            RSWHCOUNT.CommandText = sql
            RdrWHCOUNT = RSWHCOUNT.ExecuteReader

            sql = "Select count(distinct SupplierCode) COUNT,SupplierCode as WH_CODE from Schedule_Upload_Bosch_Hdr with (nolock) where SupplierCode not in (Select distinct WareHouse_Code from WareHouse_Stock_Dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' ) AND UNIT_CODE = '" & gstrUNITID & "' group by SupplierCode"
            RSWHCOUNTSTOCK.CommandText = sql
            RdrWHCOUNTSTOCK = RSWHCOUNTSTOCK.ExecuteReader

            If RdrWHCOUNT.HasRows And RdrWHCOUNTSTOCK.HasRows Then
                RdrWHCOUNT.Read()
                RdrWHCOUNTSTOCK.Read()
                If RdrWHCOUNT("Count").ToString > 0 And RdrWHCOUNT("Count").ToString = RdrWHCOUNTSTOCK("Count").ToString And RdrWHCOUNT("WH_CODE").ToString = RdrWHCOUNTSTOCK("WH_CODE").ToString Then
                    MsgBox("Stock is not defined for the Warehouse(s).So no Schedule will be proposed.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                End If
                If RdrWHCOUNT("Count").ToString > 1 Then
                    MsgBox("You have the Release File with more than 1 warehouses." & vbCrLf & "Details for these will be available in Schedule Proposal Details.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            End If
            RdrWHCOUNT.Close()
            RdrWHCOUNTSTOCK.Close()

            RSWHCOUNT.Dispose()
            RSWHCOUNT = Nothing

            RSWHCOUNTSTOCK.Dispose()
            RSWHCOUNTSTOCK = Nothing

        End If

        If ShipmentFlag = True Then
            sql = "EXEC  sp_calculatesafetystockforschedule_Bosch '" & gstrUNITID & "', '" & Me.txtCustomerCode.Text & "','" & txtConsignee.Text & "'," & "'" & trans_number & "','" & WA & "','" & sch & "','" & gstrIpaddressWinSck & "'"
            sqlCmd.CommandText = sql
            sqlCmd.CommandTimeout = 0
            sqlCmd.ExecuteNonQuery()



            Call FN_Display_BOSCH(trans_number, Darwin_FileType)
        Else
            If chkdaywisesch.Checked = True Then
                Call FN_TRANSFERDATAINCOVISINT(trans_number, Darwin_FileType)
                Darwin_FileType = "BOSCH"
            End If

            SQLTRANS = sqlCmd.Connection.BeginTransaction()
            sqlCmd.Transaction = SQLTRANS
            ISTRANS = True

            Call FN_Display_WITHOUTWH(trans_number)

            SQLTRANS.Commit()
            SQLTRANS = Nothing

        End If

        sql = "set dateformat 'dmy'"
        sqlCmd.CommandText = sql
        sqlCmd.ExecuteNonQuery()

        sql = "Select Top 1 Doc_No From ScheduleProposalcalculations Where Doc_No= '" & Val(trans_number) & "' and shipment_qty > 0 AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If Not sqlRDR.HasRows Then
            sqlRDR.Close()
            YesNo = CStr(MsgBox("No Schedule Proposed." + vbCrLf + "Do You Want To Move Files To BackUp Folder?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)))
            If YesNo = CStr(MsgBoxResult.Yes) Then Call MoveFile()
        Else
            sqlRDR.Close()
            lblMessage1.Text = "Schedule has been Uploaded Succesfully."

            Call Updt_DailyMkt(Darwin_FileType)

            If mblnfilemove = False Then
                Call MoveFile()
            End If

        End If

        sqlCmd.Dispose()
        sqlCmd = Nothing

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Sub
ERR_Renamed:

        If ISTRANS = True Then
            SQLTRANS.Rollback()
            SQLTRANS = Nothing
        End If
        If Not sqlCmd Is Nothing Then
            sqlCmd.Dispose()
            sqlCmd = Nothing
        End If

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub FN_Release_File_Upload_PKC()

        On Error GoTo ERR_Renamed

        Dim Cell_Data As String = ""
        Dim Row As Object = Nothing
        Dim i As Short = 0
        Dim Data_Row() As String = Nothing
        Dim Trans_Satus As Boolean
        Dim Upload_FileType As String = "", trans_number As String = ""
        Dim sql As String = "", Cell_Data1 As String = ""
        Dim Rev_No As Object = Nothing
        Dim Col As Short = 0

        Dim rsItems As New ClsResultSetDB
        Dim RSdt As New ClsResultSetDB
        Dim HOLIDAY As String = ""
        Dim Msg As String
        Dim sql1 As Object = Nothing
        Dim sql2 As String
        Dim rsObjectWH As New ClsResultSetDB
        Dim sqlWarehouse As String
        Dim msgWH As String = Nothing
        Dim RSconsignee As New ClsResultSetDB
        Dim RSShipmentFlag As New ClsResultSetDB
        Dim ShipmentFlag As Boolean
        Dim sheetNo As Short
        Dim YesNo As String = Nothing
        Dim WA As String
        Dim sch As Short
        Dim SftyDays, ItemCode, CustDrgNo, SftyStk, ShpgQty As Object
        Dim ShipDate As Object
        Dim wh_code
        Dim TmpRs As New ClsResultSetDB
        Dim RSAUTOMAILER As New ClsResultSetDB
        Dim sqlitem As String
        Dim sqlConsignee As String
        Dim sqlCmd As SqlCommand
        Dim sqlRDR As SqlDataReader
        Dim sqlTrans As SqlTransaction
        Dim isTrans As Boolean = False
        Dim RSWHCOUNT As SqlCommand
        Dim RSWHCOUNTSTOCK As SqlCommand
        Dim intCount As Integer = 0
        Dim RDRWHCOUNT As SqlDataReader
        Dim RDRWHCOUNTSTOCK As SqlDataReader

        Flag = 0
        HOLIDAY = ""
        msgWH = ""

        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection
        sqlCmd.CommandType = CommandType.Text

        sqlCmd.CommandText = "DELETE FROM TMPWHSTOCK WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        sqlCmd.CommandText = "DELETE FROM SCHEDULE_UPLOAD_PKC_HDR_TEMP WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        sqlCmd.CommandText = "DELETE FROM SCHEDULE_UPLOAD_PKC_DTL_TEMP WHERE UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

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

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        sql = "Select current_no + 1 as current_no from documenttype_mst where Doc_Type = 302 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            trans_number = sqlRDR("current_no").ToString
        Else
            MessageBox.Show("Document Number Not Generated.", ResolveResString(100), MessageBoxButtons.OK)
            Exit Sub
        End If
        sqlRDR.Close()

        sql = "SELECT iSnull(SHIPMENTTHRUWH,0) as SHIPMENTTHRUWH FROM CUSTOMER_MST WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            ShipmentFlag = sqlRDR("SHIPMENTTHRUWH").ToString
            sqlRDR.Close()
        Else
            MessageBox.Show("Shipment Flag Not Defined in Customer Master.", ResolveResString(100), MessageBoxButtons.OK)
            sqlRDR.Close()
            Exit Sub
        End If

        Trans_Satus = True
        If Len(Cell_Data) < 10 Then
            Col = 1
            range = Obj_EX.Cells(Row, Col)
            If Not range.Value Is Nothing Then
                Cell_Data1 = (range.Value.ToString)
            Else
                Cell_Data1 = ""
            End If
        Else
            Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
            Cell_Data1 = Trim(Data_Row(0))
        End If
        Cell_Data1 = Replace(Trim(Cell_Data1), "'", "")

        If Darwin_FileType = "PKC" Then
            Msg = ""
            sql = "Insert Into Schedule_Upload_PKC_Hdr_temp(Doc_no,CallOffNo,Customer_code,EntDt,EntBy,UpdDt,UpdBy,Unit_Code,UploadFileName)"
            sql = sql & " Values (" & trans_number & ",'" & Cell_Data1 & "','" & Trim(txtCustomerCode.Text) & "',GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "','" & gstrUNITID & "','" & txtFileName.Text & "')"

            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sqlCmd.CommandText = "set dateformat 'dmy'"
            sqlCmd.ExecuteNonQuery()

            While Len(Cell_Data) <> 0

                If Cell_Data = "PKSDF" Then
                    Col = 1 : i = 0
                    Cell_Data = ""
                    range = Obj_EX.Cells(Row, Col)
                    If Not range.Value Is Nothing Then
                        Cell_Data1 = (range.Value.ToString)
                    Else
                        Cell_Data1 = ""
                    End If
                    While Cell_Data1 <> ""
                        Cell_Data = Cell_Data & Cell_Data1 & ","
                        Col = Col + 1
                        range = Obj_EX.Cells(Row, Col)
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

                If ShipmentFlag = True Then
                    If Trim(Data_Row(6)) = "" Then
                        MsgBox("Warehouse Code is blank. File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Obj_FSO = Nothing
                        If Not Obj_EX Is Nothing Then
                            KillExcelProcess(Obj_EX)
                            Obj_EX = Nothing
                        End If
                        Exit Sub
                    End If

                    sql = "select customer_code from customer_mst " & " where cust_vendor_code = '" & Trim(Data_Row(6)) & "'" & " and dock_code = '" & Trim(Data_Row(12)) & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                    sqlCmd.CommandText = sql
                    sqlRDR = sqlCmd.ExecuteReader
                    If sqlRDR.HasRows Then
                        sqlRDR.Read()
                        consignee = sqlRDR("customer_code").ToString
                    Else
                        consignee = txtCustomerCode.Text
                    End If
                    sqlRDR.Close()
                End If

                If FN_Date_Conversion(Trim(Data_Row(10))) = "" Then
                    MsgBox("Date is blank. File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                End If

                If Trim(Data_Row(7)) = "" Then
                    MsgBox("Item Code is blank.File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If

                    Exit Sub
                End If

                If intCount = 0 Then
                    sql = " update Schedule_Upload_PKC_hdr_temp set CallOffDate = '" & FN_Date_Conversion(Trim(Data_Row(2))) & "', " & _
                        " SupplyToBuyerPlantcode = '" & Data_Row(4) & "',	SupplyFromPlantCode = '" & Data_Row(5) & "',CONSIGNEE_CODE = '" & consignee & "',  " & _
                        " SupplierCode = '" & Data_Row(6) & "',	ReferenceOrderNumber = '" & Data_Row(7) & "',	PrevInvoiceNo = '" & Data_Row(8) & "'" & _
                        " where Unit_Code = '" & gstrUNITID & "' and Doc_no = '" & trans_number & "'"
                    sqlCmd.CommandText = sql
                    sqlCmd.ExecuteNonQuery()

                    intCount = 1
                End If

                sql = "INSERT INTO SCHEDULE_UPLOAD_PKC_DTL_TEMP"
                sql = sql & "("
                sql = sql & " DOC_NO,BUYERSPARTNUMBER,CUMMQTYSTARTDATE,HORIZONENDDATE,DISPATCHQTY,SCHEDULEDATE,ENTDT,ENTBY,UPDDT,UPDBY,UNIT_CODE)"
                sql = sql & " VALUES( '" & Val(trans_number) & "', '" & Data_Row(7) & "','" & FN_Date_Conversion(Trim(Data_Row(2))) & "','" & FN_Date_Conversion(Trim(Data_Row(3))) & "','" & Data_Row(9) & "',"
                sql = sql & "'" & FN_Date_Conversion(Trim(Data_Row(10))) & " ',GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "','" & gstrUNITID & "') "

                sqlCmd.CommandText = sql
                sqlCmd.ExecuteNonQuery()

                sqlitem = "select COUNT(CUST_DRGNO) COUNT,CUST_DRGNO from custitem_mst c where c.UNIT_CODE = '" & gstrUNITID & "'" & _
                    " and c.account_code = '" & txtCustomerCode.Text & "' AND C.CUST_DRGNO = '" & Data_Row(7) & "' " & _
                    " GROUP BY CUST_DRGNO HAVING COUNT(CUST_DRGNO)  > 1"

                sqlCmd.CommandText = sqlitem
                sqlRDR = sqlCmd.ExecuteReader

                If sqlRDR.HasRows Then
                    While sqlRDR.Read
                        Msg = Msg & "'" + sqlRDR("CUST_DRGNO").ToString + "'" + vbCrLf
                    End While
                End If

                sqlRDR.Close()

                If Msg <> "" Then
                    MsgBox("Following Cust_DrgNo(s) Are Active For Multiple Items " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                End If

                sql = "select Distinct dt,Cust_Code from Calendar_mkt_Cust where dt = '" & FN_Date_Conversion(Trim(Data_Row(10))) & "' " & _
                    " AND work_flg = 1 AND UNIT_CODE = '" & gstrUNITID & "' " & _
                    " and Cust_Code = '" & txtCustomerCode.Text & "' "

                sqlCmd.CommandText = sql
                sqlRDR = sqlCmd.ExecuteReader

                If sqlRDR.HasRows Then
                    While sqlRDR.Read
                        If InStr(Replace(HOLIDAY, sqlRDR("cust_code").ToString + "---" + sqlRDR("dt").ToString, "$"), "$") = 0 Then
                            HOLIDAY = HOLIDAY & sqlRDR("cust_code").ToString + "---" + sqlRDR("dt").ToString & " "
                        End If
                    End While
                End If
                sqlRDR.Close()

                Row = Row + 1
                range = Obj_EX.Cells(Row, 1)
                If Not range.Value Is Nothing Then
                    Cell_Data = (range.Value.ToString)
                Else
                    Cell_Data = ""
                End If

            End While

        End If

        Dim countREC As Short

        sql = "SELECT Distinct BuyersPartNumber FROM Schedule_Upload_PKC_Dtl_temp D" & _
            " WHERE Unit_Code = '" & gstrUNITID & "' AND Doc_no = '" & trans_number & "'" & _
            " AND Not EXISTS (SELECT * FROM CustItem_Mst C WHERE C.UNIT_CODE = D.Unit_Code" & _
            " AND C.Cust_Drgno = D.BuyersPartNumber and c.account_code = '" & txtCustomerCode.Text & "')"

        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader

        Msg = ""
        If sqlRDR.HasRows Then
            While sqlRDR.Read
                Msg = Msg & "'" + sqlRDR("BuyersPartNumber").ToString + "'" + vbCrLf
            End While

            If Len(Trim(Msg)) > 0 Then
                MsgBox("Following Items Are Not Defined In Customer Item Master" & vbCrLf & "for Customer " & txtCustomerCode.Text & " : " & vbCrLf & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                Flag = 1
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            End If

            Trans_Satus = False

        End If
        sqlRDR.Close()

        HOLIDAY = Replace(HOLIDAY, " ", Chr(13))
        If HOLIDAY <> "" Then
            MsgBox("Following is/are not working day(s):" & vbCrLf & vbCrLf & "Consignee---Date" & vbCrLf & HOLIDAY & vbCrLf)
            Flag = 1
        End If

        If ShipmentFlag = True Then
            sql = "select distinct suppliercode from schedule_upload_PKC_HDR_temp t where Unit_Code = '" & gstrUNITID & "' and doc_no = '" & trans_number & "'" & _
                " and Not exists (select top 1 1 from scheduleparameter_mst where UNIT_CODE = t.Unit_Code and WH_code = t.SupplierCode)"

            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader

            If sqlRDR.HasRows Then
                sqlRDR.Read()
                msgWH = msgWH & "  '" + sqlRDR("SupplierCode").ToString + "'  "
                Flag = 1
            End If

            sqlRDR.Close()

            If msgWH <> "" Then
                MsgBox("WRONG WAREHOUSE: " & msgWH, MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
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

        If Flag = 1 Then
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Exit Sub
        Else
            If ShipmentFlag = False And chkDlyPullQty.Checked = True Then
                MsgBox("Shipment For This Customer Is Not Through Warehouse, " & vbCrLf & "So, You Can't Use Daily Pull Qty For Safety Stock Calculations.", vbInformation, ResolveResString(100))
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            Else
                Me.txtDocNo.Text = trans_number
            End If
        End If

        If Flag = 0 Then
            sqlTrans = sqlCmd.Connection.BeginTransaction()
            sqlCmd.Transaction = sqlTrans
            isTrans = True

            sql = "INSERT INTO schedule_upload_PKC_hdr SELECT * FROM schedule_upload_PKC_hdr_TEMP where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "INSERT INTO schedule_upload_PKC_DTL SELECT * FROM schedule_upload_PKC_DTL_TEMP where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "update documenttype_mst set current_no = " & trans_number & " where Doc_Type = 302 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sqlTrans.Commit()
            isTrans = False
        End If

        If ShipmentFlag = True Then
            sql = "EXEC SP_CALCULATESAFETYSTOCKFORSCHEDULE_PKC '" & gstrUNITID & "', '" & Me.txtCustomerCode.Text & "'" & ",'" & trans_number & "'," & "'" & WA & "','" & sch & "','" & gstrIpaddressWinSck & "'"
            sqlCmd.CommandText = sql
            sqlCmd.CommandTimeout = 0
            sqlCmd.ExecuteNonQuery()

            sql = "Select top 1 1 from schedule_upload_PKC_hdr h inner join warehouse_stock_dtl w" & _
                " on h.Unit_Code = w.UNIT_CODE and h.Customer_Code = w.Customer_Code and h.SupplierCode = w.WareHouse_Code " & _
                " where  h.UNIT_CODE = '" & gstrUNITID & "' and h.Doc_no = '" & trans_number & "'"

            RSWHCOUNTSTOCK = New SqlCommand
            RSWHCOUNTSTOCK.Connection = SqlConnectionclass.GetConnection
            RSWHCOUNTSTOCK.CommandType = CommandType.Text

            RSWHCOUNTSTOCK.CommandText = sql
            RDRWHCOUNTSTOCK = RSWHCOUNTSTOCK.ExecuteReader

            If Not RDRWHCOUNTSTOCK.HasRows Then
                MsgBox("Stock is not defined for the Warehouse(s).So no Schedule will be proposed.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            End If

            RDRWHCOUNTSTOCK.Close()
            RSWHCOUNTSTOCK.Dispose()
            RSWHCOUNTSTOCK = Nothing

        End If

        If ShipmentFlag = True Then
            Call FN_Display_PKC(trans_number, Darwin_FileType)
        Else
            Call FN_Display_WITHOUTWH(trans_number)
        End If

        sqlCmd.CommandText = "set dateformat 'dmy'"
        sqlCmd.ExecuteNonQuery()

        sql = "Select Top 1 Doc_No From ScheduleProposalcalculations Where Doc_No= '" & Val(trans_number) & "' and shipment_qty > 0 AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader

        If Not sqlRDR.HasRows Then
            sqlRDR.Close()
            YesNo = CStr(MsgBox("No Schedule Proposed." + vbCrLf + "Do You Want To Move Files To BackUp Folder?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)))

            If YesNo = CStr(MsgBoxResult.Yes) Then
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                lblMessage1.Text = "No Schedule Proposed."
                Call MoveFile()
            End If
        Else
            sqlRDR.Close()
            lblMessage1.Text = "Schedule has been Uploaded Succesfully."

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            sqlCmd.Dispose()
            sqlCmd = Nothing

            Call Updt_DailyMkt(Darwin_FileType)

            If mblnfilemove = False Then
                Call MoveFile()
            End If

        End If
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        TmpRs = Nothing
        Exit Sub
ERR_Renamed:

        If isTrans = True Then
            sqlTrans.Rollback()
            sqlTrans = Nothing
        End If

        If Not sqlCmd Is Nothing Then
            sqlCmd.Dispose()
            sqlCmd = Nothing
        End If

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)


    End Sub

    Private Sub FN_Release_File_Upload_EDIFACT()
        On Error GoTo ERR_Renamed

        Dim Cell_Data As String = ""
        Dim Row As Object = Nothing
        Dim i As Short = 0
        Dim Data_Row() As String = Nothing
        Dim Trans_Satus As Boolean
        Dim Upload_FileType As String = "", trans_number As String = ""
        Dim sql As String = "", Cell_Data1 As String = ""
        Dim Rev_No As Object = Nothing
        Dim Col As Short = 0
        Dim HOLIDAY As String = ""
        Dim Msg As String
        Dim sql1 As Object = Nothing
        Dim sql2 As String
        Dim msgWH As String = Nothing
        Dim ShipmentFlag As Boolean
        Dim sheetNo As Short
        Dim YesNo As String = Nothing
        Dim WA As String
        Dim sch As Integer
        Dim SftyDays, ItemCode, CustDrgNo, SftyStk, ShpgQty As Object
        Dim ShipDate As Object
        Dim wh_code
        Dim RSWHCOUNT As SqlCommand
        Dim RSWHCOUNTSTOCK As SqlCommand
        Dim RdrWHCOUNT As SqlDataReader
        Dim RdrWHCOUNTSTOCK As SqlDataReader

        Dim sqlCmd As SqlCommand
        Dim sqlRDR As SqlDataReader
        Dim SQLTRANS As SqlTransaction
        Dim ISTRANS As Boolean

        Flag = 0
        HOLIDAY = ""
        msgWH = ""

        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection
        sqlCmd.CommandType = CommandType.Text

        sqlCmd.CommandText = "DELETE FROM TMPWHSTOCK WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        sqlCmd.CommandText = "DELETE FROM SCHEDULE_UPLOAD_DARWINEDIFACT_HDR_TEMP WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        sqlCmd.CommandText = "DELETE FROM SCHEDULE_UPLOAD_DARWINEDIFACT_DTL_TEMP WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()


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

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        sql = "Select current_no + 1 as current_no from documenttype_mst where Doc_Type = 302 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            trans_number = sqlRDR("current_no").ToString
        Else
            MessageBox.Show("Document Number Not Generated.", ResolveResString(100), MessageBoxButtons.OK)
            Exit Sub
        End If
        sqlRDR.Close()

        sql = "SELECT iSnull(SHIPMENTTHRUWH,0) as SHIPMENTTHRUWH FROM CUSTOMER_MST WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            ShipmentFlag = sqlRDR("SHIPMENTTHRUWH").ToString
            sqlRDR.Close()
        Else
            MessageBox.Show("Shipment Flag Not Defined in Customer Master.", ResolveResString(100), MessageBoxButtons.OK)
            sqlRDR.Close()
            Exit Sub
        End If

        Trans_Satus = True
        If Len(Cell_Data) < 10 Then
            Col = 1
            range = Obj_EX.Cells(Row, Col)
            If Not range.Value Is Nothing Then
                Cell_Data1 = (range.Value.ToString)
            Else
                Cell_Data1 = ""
            End If
        Else
            Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
            Cell_Data1 = Trim(Data_Row(0))
        End If
        Cell_Data1 = Replace(Trim(Cell_Data1), "'", "")

        Msg = ""
        If Len(Cell_Data) < 10 Then
            Col = 1 : i = 0
            Cell_Data = ""
            If Not range.Value Is Nothing Then
                Cell_Data1 = (range.Value.ToString)
            Else
                Cell_Data1 = ""
            End If
            While Cell_Data1 <> ""
                Cell_Data = Cell_Data & Cell_Data1 & ","
                Col = Col + 1
                range = Obj_EX.Cells(Row, Col)
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

        sql1 = " Insert Into Schedule_Upload_DarwinEDIFACT_Hdr_temp(Doc_No,Doc_Type,Cust_Code," & _
            "Consignee_Code,Plant_c,Upload_File_Name,Upload_File_Type,SenderID," & _
            "RecipientID,Receipt_Dt,Receipt_Time,Control_no,Test_Indicator,Msg_code," & _
            "Msg_Name,Msg_Number,Msg_Version,Upload_DtQualifier,Upload_Dt," & _
            "Upload_DtFormatQualifier,Start_DtQualifier,Start_Dt,Start_DtFormatQualifier,End_DtQualifier,End_Dt,End_DtFormatQualifier," & _
            "Party_Qualifier1,Party_ID1,Agency_code1,Party_Qualifier2,PARTY_ID2," & _
            "Agency_code2,Process_Indicator,Party_Qualifier3,Party_ID3,Agency_code3," & _
            "Ent_Dt,Upd_Dt,Ent_UId,Upd_UId,UNIT_CODE) " & _
            " Values (" & trans_number & ",302,'" & Trim(txtCustomerCode.Text) & "'," & _
            " '','" & Trim(txtUnitCode.Text) & "'," & _
            "'" & Trim(txtFileName.Text) & "','" & Upload_FileType & "'," & _
            " '" & Trim(Data_Row(1)) & "','" & Trim(Data_Row(2)) & "'," & _
            " '" & Format(CDate(FN_Date_Conversion_edifact(Trim(Data_Row(3)))), "dd MMM yyyy") & "'," & _
            " " & Val(Data_Row(4)) & ",'" & Trim(Data_Row(5)) & "'," & Val(Data_Row(6)) & "," & " '" & Trim(Data_Row(7)) & "', '" & Trim(Data_Row(8)) & "','" & Val(Data_Row(9)) & "'," & _
            " '" & Val(Data_Row(10)) & "', '" & Val(Data_Row(11)) & "','" & Format(CDate(FN_Date_Conversion_edifact(Trim(Data_Row(12)))), "dd MMM yyyy") & "'," & " '" & Trim(Data_Row(13)) & "', '" & Trim(Data_Row(14)) & "'," & _
            " '" & Format(CDate(FN_Date_Conversion_edifact(Trim(Data_Row(15)))), "dd MMM yyyy") & "'," & _
            " '" & Trim(Data_Row(16)) & "', '" & Trim(Data_Row(17)) & "'," & _
            " '" & Format(CDate(FN_Date_Conversion_edifact(Trim(Data_Row(18)))), "dd MMM yyyy") & "'," & _
            " '" & Trim(Data_Row(19)) & "', '" & Trim(Data_Row(20)) & "','" & Trim(Data_Row(21)) & "'," & _
            " '" & Trim(Data_Row(22)) & "', '" & Trim(Data_Row(23)) & "','" & Trim(Data_Row(24)) & "'," & _
            " '" & Trim(Data_Row(25)) & "', '" & Trim(Data_Row(26)) & "','" & Trim(Data_Row(27)) & "'," & _
            " '" & Trim(Data_Row(28)) & "', '" & Trim(Data_Row(29)) & "',getDate(),getDate(), " & _
            " '" & mP_User & "','" & mP_User & "','" & gstrUNITID & "') "

        sqlCmd.CommandText = sql1
        sqlCmd.ExecuteNonQuery()

        While Len(Cell_Data) <> 0

            If Len(Cell_Data) < 10 Then
                Col = 1 : i = 0
                Cell_Data = ""
                range = Obj_EX.Cells(Row, Col)
                If Not range.Value Is Nothing Then
                    Cell_Data1 = Trim(range.Value.ToString)
                Else
                    Cell_Data1 = ""
                End If

                While Cell_Data1 <> ""
                    Cell_Data = Cell_Data & Cell_Data1 & ","
                    Col = Col + 1
                    range = Obj_EX.Cells(Row, Col)
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
                MsgBox("Item Code is blank.File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            End If

            If FN_Date_Conversion_edifact(Trim(Data_Row(3))) = "" Then
                MsgBox("Date is blank. File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If

                Exit Sub

                'ElseIf CDate(FN_Date_Conversion_edifact(Trim(Data_Row(3)))) > CDate(IIf(FN_Date_Conversion_edifact(Trim(Data_Row(46))) = "", "01/01/1900", FN_Date_Conversion_edifact(Trim(Data_Row(46))))) Then
                '    If Trim(Data_Row(40)) <> "" And Trim(Data_Row(41)) <> "" Then
                '        MsgBox("Schedule Date " + FN_Date_Conversion_edifact(Trim(Data_Row(46))) + " Should Be Greater Than" + vbCrLf + "Transmission Date " + FN_Date_Conversion_edifact(Trim(Data_Row(3))) + " of Release File" + vbCrLf + "File Can't Be Uploaded", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                '        Obj_FSO = Nothing
                '        If Not Obj_EX Is Nothing Then
                '            KillExcelProcess(Obj_EX)
                '            Obj_EX = Nothing
                '        End If

                '        Exit Sub
                '    End If
            End If

            Dim STRCONS As String
            sql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST WHERE CUST_VENDOR_CODE = '" & Trim(Data_Row(21)) & "' AND DOCK_CODE = '" & Trim(Data_Row(28)) & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader
            If sqlRDR.HasRows Then
                sqlRDR.Read()
                STRCONS = sqlRDR("CUSTOMER_CODE").ToString
            Else
                STRCONS = txtCustomerCode.Text
            End If

            sqlRDR.Close()

            sql = " Insert into Schedule_Upload_DarwinEDIFACT_dtl_TEMP(Doc_No,Doc_Type, Cust_Code,Consignee_Code,Item_Code,Item_Type,ProductID_Qualifier," & _
                " Item_Number,Item_NumberType,Location_Qualifier,Location_ID,Ref_Qualifier,Ref_ID, DelPlan_Status,Frequency,Dispatch_Pattern," & _
                " Quantity_Qualifier,Quantity,UOM," & " DelDT_Qualifier,Delivery_DT,DelDT_FormatQualifier,Ent_Dt,Upd_Dt,Ent_UId," & _
                " Upd_UId,UNIT_CODE) " & _
                " Values (" & trans_number & ",302,'" & Trim(txtCustomerCode.Text) & "'," & " '" & Trim(STRCONS) & "','" & Trim(Data_Row(30)) & "'," & _
                " '" & Trim(Data_Row(31)) & "', " & " '" & Trim(Data_Row(32)) & "','" & Trim(Data_Row(33)) & "','" & Trim(Data_Row(34)) & "', " & _
                " '" & Trim(Data_Row(35)) & "','" & Trim(Data_Row(36)) & "','" & Trim(Data_Row(37)) & "', " & _
                " '" & Trim(Data_Row(38)) & "','" & Trim(Data_Row(39)) & "','" & Trim(Data_Row(40)) & "', " & _
                " '" & Trim(Data_Row(41)) & "','" & Trim(Data_Row(42)) & "'," & Val(Data_Row(43)) & ", " & _
                " '" & Trim(Data_Row(44)) & "','" & Trim(Data_Row(45)) & "','" & Trim(Data_Row(46)) & "', " & _
                " '" & Trim(Data_Row(47)) & "',getDate(),getDate() ," & " '" & mP_User & "','" & mP_User & "','" & gstrUNITID & "')"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            Row = Row + 1
            range = Obj_EX.Cells(Row, 1)
            If Not range.Value Is Nothing Then
                Cell_Data = (range.Value.ToString)
            Else
                Cell_Data = ""
            End If
        End While

        sql = "select cust_drgno FROM CUSTITEM_MST " & _
            " WHERE CUST_DRGNO = '" & Trim(Data_Row(30)) & "'" & " AND active = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "' AND"
        sql = sql & " account_code = '" & Me.txtCustomerCode.Text & "' GROUP BY Cust_Drgno HAVING COUNT(CUST_DRGNO)  > 1"

        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader

        If sqlRDR.HasRows Then
            sqlRDR.Read()
            Msg = Msg & "'" + sqlRDR("CUST_DRGNO").ToString + "'" + vbCrLf
        End If
        sqlRDR.Close()

        sql = "select distinct d.ITEM_CODE " & _
                " from Schedule_Upload_DarwinEDIFACT_dtl_temp d,Schedule_Upload_DarwinEDIFACT_hdr_temp h " & _
                " Where h.cust_code = '" & Me.txtCustomerCode.Text & "'" & _
                " and d.doc_no = h.doc_no AND D.UNIT_CODE = H.UNIT_CODE and h.doc_no=" & trans_number & " AND d.UNIT_CODE = '" & gstrUNITID & "'" & _
                " and ltrim(rtrim(d.ITEM_CODE)) " & _
                " not in (select cust_drgno from custitem_mst where active = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "'"
        If ShipmentFlag = True Then
            sql = sql + " and account_code in (SELECT DISTINCT CONSIGNEE_CODE FROM SCHEDULE_UPLOAD_DARWINEDIFACT_DTL_temp WHERE DOC_NO = " & trans_number & " AND UNIT_CODE = '" & gstrUNITID & "'))"
        Else
            sql = sql & " and account_code = '" & Me.txtCustomerCode.Text & "')"
        End If

        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        Msg = ""

        If sqlRDR.HasRows Then
            While sqlRDR.Read
                Msg = Msg & "'" + sqlRDR("ITEM_CODE").ToString + "'" + vbCrLf
            End While

            If Len(Trim(Msg)) > 0 Then
                MsgBox("Following Items Are Not Defined In The System" & vbCrLf & "for Customer " & txtCustomerCode.Text & " : " & vbCrLf & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                Flag = 1
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            End If

            Trans_Satus = False

        End If
        sqlRDR.Close()

        HOLIDAY = Replace(HOLIDAY, " ", Chr(13))
        If HOLIDAY <> "" Then
            MsgBox("Following is/are not working day(s):" & vbCrLf & vbCrLf & "Consignee---Date" & vbCrLf & HOLIDAY & vbCrLf)
            Flag = 1
        End If

        If ShipmentFlag = True Then
            sql = "select distinct h.PARTY_ID1 from schedule_upload_darwinedifact_hdr_temp H," & _
                    " scheduleparameter_mst s where h.PARTY_ID1  not in(select wh_code " & _
                    " from scheduleparameter_mst s where s.customer_code =  '" & Me.txtCustomerCode.Text & "' AND S.UNIT_CODE = '" & gstrUNITID & "')" & _
                    " and cust_code = '" & Me.txtCustomerCode.Text & "' and doc_no = " & trans_number & " AND S.UNIT_CODE = H.UNIT_CODE AND S.UNIT_CODE = '" & gstrUNITID & "' "
            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader

            If sqlRDR.HasRows Then
                While sqlRDR.Read
                    msgWH = msgWH & "  '" + sqlRDR("PARTY_ID1").ToString + "'  "
                    Flag = 1
                End While
            End If
            sqlRDR.Close()

            If msgWH <> "" Then
                MsgBox("WRONG WAREHOUSE: " & msgWH, MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
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

        If Flag = 1 Then
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Exit Sub
        Else
            If ShipmentFlag = False And chkDlyPullQty.Checked = True Then
                MsgBox("Shipment For This Customer Is Not Through Warehouse, " & vbCrLf & "So, You Can't Use Daily Pull Qty For Safety Stock Calculations.", vbInformation, ResolveResString(100))
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            Else
                Me.txtDocNo.Text = trans_number
            End If
        End If

        If Flag = 0 Then
            SQLTRANS = sqlCmd.Connection.BeginTransaction()
            sqlCmd.Transaction = SQLTRANS
            ISTRANS = True

            sql = "INSERT INTO schedule_upload_DARWINEDIFACT_hdr SELECT * FROM schedule_upload_DARWINEDIFACT_hdr_TEMP where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "INSERT INTO schedule_upload_DARWINEDIFACT_DTL SELECT * FROM schedule_upload_DARWINEDIFACT_DTL_TEMP where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "update documenttype_mst set current_no = " & trans_number & " where Doc_Type = 302 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            SQLTRANS.Commit()
            ISTRANS = False
        End If

        If ShipmentFlag = True Then
            RSWHCOUNT = New SqlCommand
            RSWHCOUNT.Connection = SqlConnectionclass.GetConnection

            RSWHCOUNTSTOCK = New SqlCommand
            RSWHCOUNTSTOCK.Connection = SqlConnectionclass.GetConnection

            sql = "Select count(distinct PARTY_ID1) COUNT,PARTY_ID1 as WH_CODE from SCHEDULE_UPLOAD_DARWINEDIFACT_HDR with (nolock) where doc_no='" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "' group by PARTY_ID1"
            RSWHCOUNT.CommandText = sql
            RdrWHCOUNT = RSWHCOUNT.ExecuteReader

            sql = "Select count(distinct PARTY_ID1) COUNT,PARTY_ID1 as WH_CODE from SCHEDULE_UPLOAD_DARWINEDIFACT_HDR with (nolock) where PARTY_ID1 not in (Select distinct WareHouse_Code from WareHouse_Stock_Dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' ) AND UNIT_CODE = '" & gstrUNITID & "' group by PARTY_ID1"
            RSWHCOUNTSTOCK.CommandText = sql
            RdrWHCOUNTSTOCK = RSWHCOUNTSTOCK.ExecuteReader

            If RdrWHCOUNT.HasRows And RdrWHCOUNTSTOCK.HasRows Then
                If RdrWHCOUNT("Count").ToString > 0 And RdrWHCOUNT("Count").ToString = RdrWHCOUNTSTOCK("Count").ToString And RdrWHCOUNT("WH_CODE").ToString = RdrWHCOUNTSTOCK("WH_CODE").ToString Then
                    MsgBox("Stock is not defined for the Warehouse(s).So no Schedule will be proposed.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                End If
                If RdrWHCOUNT("Count").ToString > 1 Then
                    MsgBox("You have the Release File with more than 1 warehouses." & vbCrLf & "Details for these will be available in Schedule Proposal Details.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            End If
            RdrWHCOUNT.Close()
            RdrWHCOUNTSTOCK.Close()

            RSWHCOUNT.Dispose()
            RSWHCOUNT = Nothing

            RSWHCOUNTSTOCK.Dispose()
            RSWHCOUNTSTOCK = Nothing

        End If

        If ShipmentFlag = True Then
            sql = "EXEC  SP_CALCULATESAFETYSTOCKFORSCHEDULE_EDIFACT '" & gstrUNITID & "', '" & Me.txtCustomerCode.Text & "','" & txtConsignee.Text & "'," & "'" & trans_number & "','" & WA & "','" & sch & "','" & gstrIpaddressWinSck & "'"
            sqlCmd.CommandText = sql
            sqlCmd.CommandTimeout = 0
            sqlCmd.ExecuteNonQuery()

            If chkdaywisesch.Checked = True Then
                Call FN_TRANSFERDATAINCOVISINT(trans_number, Darwin_FileType)
                Darwin_FileType = "COVISINT"
            End If

            Call FN_Display(trans_number, Darwin_FileType)
        Else
            If chkdaywisesch.Checked = True Then
                Call FN_TRANSFERDATAINCOVISINT(trans_number, Darwin_FileType)
                Darwin_FileType = "COVISINT"
            End If

            SQLTRANS = sqlCmd.Connection.BeginTransaction()
            sqlCmd.Transaction = SQLTRANS
            ISTRANS = True

            Call FN_Display_WITHOUTWH(trans_number)

            SQLTRANS.Commit()
            SQLTRANS = Nothing

        End If

        sql = "set dateformat 'dmy'"
        sqlCmd.CommandText = sql
        sqlCmd.ExecuteNonQuery()

        sql = "Select Top 1 Doc_No From ScheduleProposalcalculations Where Doc_No= '" & Val(trans_number) & "' and shipment_qty > 0 AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If Not sqlRDR.HasRows Then
            sqlRDR.Close()
            YesNo = CStr(MsgBox("No Schedule Proposed." + vbCrLf + "Do You Want To Move Files To BackUp Folder?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)))
            If YesNo = CStr(MsgBoxResult.Yes) Then Call MoveFile()
        Else
            sqlRDR.Close()
            lblMessage1.Text = "Schedule has been Uploaded Succesfully."

            Call Updt_DailyMkt(Darwin_FileType)

            If mblnfilemove = False Then
                Call MoveFile()
            End If

        End If

        sqlCmd.Dispose()
        sqlCmd = Nothing

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Sub
ERR_Renamed:

        If ISTRANS = True Then
            SQLTRANS.Rollback()
            SQLTRANS = Nothing
        End If
        If Not sqlCmd Is Nothing Then
            sqlCmd.Dispose()
            sqlCmd = Nothing
        End If

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub FN_Release_File_Upload()
        On Error GoTo ERR_Renamed
        Dim Cell_Data As String = ""
        Dim Data_Row() As String = Nothing
        Dim Cell_Data1 As String = ""
        Dim sqlCmd As SqlCommand
        Dim sqlRDR As SqlDataReader
        Dim Sql As String = String.Empty

        Obj_EX = New Excel.Application
        Obj_EX.Workbooks.Open(Trim(Me.txtFileName.Text))

        range = Obj_EX.Cells(1, 1)
        If Not range.Value Is Nothing Then
            Cell_Data = (range.Value.ToString)
        Else
            Cell_Data = ""
        End If

        If Len(Cell_Data) = 0 Then

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        If Len(Cell_Data) < 10 Then
            range = Obj_EX.Cells(1, 1)
            If Not range.Value Is Nothing Then
                Cell_Data1 = (range.Value.ToString)
            Else
                Cell_Data1 = ""
            End If
        Else
            Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
            Cell_Data1 = Trim(Data_Row(0))
        End If
        Cell_Data1 = Replace(Trim(Cell_Data1), "'", "")

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection
        sqlCmd.CommandType = CommandType.Text

        Sql = "SELECT KEY2 FROM LISTS WHERE KEY1='CDPFILETYPE' AND DESCR='" & Cell_Data1 & "' and UNIT_CODE='" & gStrUnitId & "'"
        sqlCmd.CommandText = Sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            While sqlRDR.Read
                Darwin_FileType = sqlRDR("KEY2").ToString
            End While
        End If

        If Darwin_FileType = "VDA" Then
            Call FN_Release_File_Upload_VDA()
        ElseIf Darwin_FileType = "EDIFACT" Then
            Call FN_Release_File_Upload_EDIFACT()
        ElseIf Darwin_FileType = "COVISINT" Then
            Call FN_Release_File_Upload_COVISINT()
        ElseIf Darwin_FileType = "PKC" Then
            Call FN_Release_File_Upload_PKC()
        ElseIf Darwin_FileType = "BOSCH" Then
            Call FN_Release_File_Upload_Bosch()
        Else
            MsgBox("Wrong File format.", MsgBoxStyle.OkOnly, ResolveResString(100))
            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End If
        'If Cell_Data1 = "4905" Then
        '    Darwin_FileType = "VDA"
        '    Call FN_Release_File_Upload_VDA()
        'ElseIf Cell_Data1 = "DELFOR" Then
        '    Darwin_FileType = "EDIFACT"
        '    Call FN_Release_File_Upload_EDIFACT()
        'ElseIf LTrim(UCase(VB.Left(Cell_Data1, 8))) = "COVISINT" Then
        '    Darwin_FileType = "COVISINT"
        '    Call FN_Release_File_Upload_COVISINT()
        'ElseIf Cell_Data1 = "PKSDF" Then
        '    Darwin_FileType = "PKC"
        '    Call FN_Release_File_Upload_PKC()
        'ElseIf Cell_Data1 = "0097359171" Or Cell_Data1 = "MX02MO37" Or Cell_Data1 = "MX02AZ79" Or Cell_Data1 = "HN02GV72" Or Cell_Data1 = "HN02NC72" Or Cell_Data1 = "MX01MRR2" Or Cell_Data1 = "MX02MTR2" Or Cell_Data1 = "MX02NOV2" Or Cell_Data1 = "TORRES" Or Cell_Data1 = "UNITECH" Or Cell_Data1 = "UTA01GM71" Or Cell_Data1 = "CQWOFE" Then
        '    Darwin_FileType = "BOSCH"
        '    Call FN_Release_File_Upload_Bosch()
        'ElseIf Cell_Data1 = "0941A430039874" Then
        '    Darwin_FileType = "BOSCH"
        '    Call FN_Release_File_Upload_Bosch()
        'Else
        '    MsgBox("Wrong File format.", MsgBoxStyle.OkOnly, ResolveResString(100))
        '    Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        '    Exit Sub
        'End If

        Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        Exit Sub
ERR_Renamed:
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub FN_Release_File_Upload_VDA()
        On Error GoTo ERR_Renamed
        Dim Cell_Data As String = ""
        Dim Row As Object = Nothing
        Dim i As Short = 0
        Dim Data_Row() As String = Nothing
        Dim Trans_Satus As Boolean
        Dim Upload_FileType As String = "", trans_number As String = ""
        Dim sql As String = "", Cell_Data1 As String = ""
        Dim Rev_No As Object = Nothing
        Dim Col As Short = 0

        Dim rsItems As New ClsResultSetDB
        Dim RSdt As New ClsResultSetDB
        Dim HOLIDAY As String = ""
        Dim Msg As String
        Dim sql1 As Object = Nothing
        Dim sql2 As String
        Dim rsObjectWH As New ClsResultSetDB
        Dim sqlWarehouse As String
        Dim msgWH As String = Nothing
        Dim RSconsignee As New ClsResultSetDB
        Dim RSShipmentFlag As New ClsResultSetDB
        Dim ShipmentFlag As Boolean
        Dim sheetNo As Short
        Dim YesNo As String = Nothing
        Dim WA As String
        Dim sch As Short
        Dim SftyDays, ItemCode, CustDrgNo, SftyStk, ShpgQty As Object
        Dim ShipDate As Object
        Dim wh_code
        Dim TmpRs As New ClsResultSetDB
        Dim RSAUTOMAILER As New ClsResultSetDB
        Dim sqlitem As String
        Dim sqlConsignee As String
        Dim sqlCmd As SqlCommand
        Dim sqlRDR As SqlDataReader
        Dim sqlTrans As SqlTransaction
        Dim isTrans As Boolean = False
        Dim RSWHCOUNT As SqlCommand
        Dim RSWHCOUNTSTOCK As SqlCommand

        Dim RDRWHCOUNT As SqlDataReader
        Dim RDRWHCOUNTSTOCK As SqlDataReader

        Flag = 0
        HOLIDAY = ""
        msgWH = ""

        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection
        sqlCmd.CommandType = CommandType.Text

        sqlCmd.CommandText = "DELETE FROM TMPWHSTOCK WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        sqlCmd.CommandText = "DELETE FROM SCHEDULE_UPLOAD_DARWIN_HDR_TEMP WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

        sqlCmd.CommandText = "DELETE FROM SCHEDULE_UPLOAD_DARWIN_DTL_TEMP WHERE UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.ExecuteNonQuery()

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

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        sql = "Select current_no + 1 as current_no from documenttype_mst where Doc_Type = 302 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            trans_number = sqlRDR("current_no").ToString
        Else
            MessageBox.Show("Document Number Not Generated.", ResolveResString(100), MessageBoxButtons.OK)
            Exit Sub
        End If
        sqlRDR.Close()

        sql = "SELECT iSnull(SHIPMENTTHRUWH,0) as SHIPMENTTHRUWH FROM CUSTOMER_MST WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        If sqlRDR.HasRows Then
            sqlRDR.Read()
            ShipmentFlag = sqlRDR("SHIPMENTTHRUWH").ToString
            sqlRDR.Close()
        Else
            MessageBox.Show("Shipment Flag Not Defined in Customer Master.", ResolveResString(100), MessageBoxButtons.OK)
            sqlRDR.Close()
            Exit Sub
        End If

        Trans_Satus = True
        If Len(Cell_Data) < 10 Then
            Col = 1
            range = Obj_EX.Cells(Row, Col)
            If Not range.Value Is Nothing Then
                Cell_Data1 = (range.Value.ToString)
            Else
                Cell_Data1 = ""
            End If
        Else
            Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
            Cell_Data1 = Trim(Data_Row(0))
        End If
        Cell_Data1 = Replace(Trim(Cell_Data1), "'", "")

        If Darwin_FileType = "VDA" Then
            Msg = ""
            sql = "Insert Into Schedule_Upload_Darwin_Hdr_temp(Doc_No,Doc_Type,Cust_Code,consignee_code,Plant_c,Upload_File_Name,Upload_File_Type,Ent_Dt,Upd_Dt,Ent_UId,Upd_UId,UNIT_CODE)"
            sql = sql & " Values (" & trans_number & ",302,'" & Trim(txtCustomerCode.Text) & "','" & Trim(txtConsignee.Text) & "','" & Trim(txtUnitCode.Text) & "','" & Trim(txtFileName.Text) & "','" & Upload_FileType & "',GETDATE(),GETDATE(),'" & mP_User & "','" & mP_User & "','" & gstrUNITID & "')"

            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sqlCmd.CommandText = "set dateformat 'dmy'"
            sqlCmd.ExecuteNonQuery()

            While Len(Cell_Data) <> 0

                If Len(Cell_Data) < 10 Then
                    Col = 1 : i = 0
                    Cell_Data = ""
                    range = Obj_EX.Cells(Row, Col)
                    If Not range.Value Is Nothing Then
                        Cell_Data1 = (range.Value.ToString)
                    Else
                        Cell_Data1 = ""
                    End If
                    While Cell_Data1 <> ""
                        Cell_Data = Cell_Data & Cell_Data1 & ","
                        Col = Col + 1
                        range = Obj_EX.Cells(Row, Col)
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

                If ShipmentFlag = True Then
                    If Trim(Data_Row(4)) = "" Then
                        MsgBox("Warehouse Code is blank.File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Obj_FSO = Nothing
                        If Not Obj_EX Is Nothing Then
                            KillExcelProcess(Obj_EX)
                            Obj_EX = Nothing
                        End If
                        Exit Sub
                    End If

                    sql = "select customer_code from customer_mst " & " where cust_vendor_code = '" & Trim(Data_Row(4)) & "'" & " and dock_code = '" & Trim(Data_Row(12)) & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                    sqlCmd.CommandText = sql
                    sqlRDR = sqlCmd.ExecuteReader
                    If sqlRDR.HasRows Then
                        sqlRDR.Read()
                        consignee = sqlRDR("customer_code").ToString
                    Else
                        consignee = txtCustomerCode.Text
                    End If
                    sqlRDR.Close()
                End If

                If FN_Date_Conversion(Trim(Data_Row(7))) = "" Then
                    MsgBox("Date is blank. File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                Else
                    If Data_Row(37) = "444444" Or Data_Row(37) = "555555" Then
                        Data_Row(37) = Data_Row(7)
                    End If
                End If


                If Trim(Data_Row(17)) = "" Then
                    MsgBox("Item Code is blank.File can't be uploaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If

                    Exit Sub
                End If

                sql = "Insert into Schedule_Upload_Darwin_Dtl_temp"
                sql = sql & "("
                sql = sql & " Doc_No,Doc_Type,"
                sql = sql & " GI_RecName,GI_Version,GI_Cust_Code,GI_Vend_Code,GI_TransNum_Old,GI_TransNum_New,GI_TransDate,"
                sql = sql & " GI_RestDT_Deli_Status_No , GI_Seg_Padding,"

                sql = sql & " GDRI_RecName,GDRI_Version,GDRI_Factory_Code,GDRI_Del_Req_No_New,GDRI_Del_Req_Dt_New,GDRI_Del_Req_No_Old,"
                sql = sql & " GDRI_Del_Req_DT_Old,GDRI_Cust_Article_No,GDRI_Vend_Article_No,GDRI_Order_No,GDRI_Unloading_Area,"
                sql = sql & " GDRI_Iden_Of_Cust,GDRI_UOM,GDRI_Del_Interval,GDRI_Pro_Release,GDRI_Mat_Release,GDRI_PurposeKey,"
                sql = sql & " GDRI_Addi_Key,GDRI_Place_Storage,GDRI_Seg_Padding,"

                sql = sql & "DDRD_RecName,DDRD_Version,DDRD_Capture_Dt,DDRD_Delivery_Note_No,DDRD_Delivery_Note_Dt,DDRD_QTY_Last_Receipt,"
                sql = sql & "DDRD_Deli_Status_No,DDRD_Req_Dt1,DDRD_Req_Qty1,DDRD_Req_Dt2,DDRD_Req_Qty2,DDRD_Req_Dt3,DDRD_Req_Qty3,DDRD_Req_Dt4,"
                sql = sql & "DDRD_Req_Qty4,DDRD_Req_Dt5,DDRD_Req_Qty5,DDRD_Seg_Padding,"

                sql = sql & "ADRD_RecName,ADRD_Version,ADRD_Req_Dt6,ADRD_Req_Qty6 ,ADRD_Req_Dt7 ,ADRD_Req_Qty7 ,"
                sql = sql & "ADRD_Req_Dt8,ADRD_Req_Qty8,ADRD_Req_Dt9 ,ADRD_Req_Qty9,ADRD_Req_Dt10 ,ADRD_Req_Qty10 ,ADRD_Req_Dt11 ,"
                sql = sql & "ADRD_Req_Qty11,ADRD_Req_Dt12,ADRD_Req_Qty12 ,"
                sql = sql & "ADRD_Req_Dt13,ADRD_Req_Qty13,ADRD_SegPadding,"

                sql = sql & "PI_RecName,PI_Version,PI_Prod_Relese_Start_Dt,PI_Prod_Relese_Finish_Dt  ,"
                sql = sql & "PI_Prod_Relese_Deli_Status,PI_Mat_Relese_Start_Dt,PI_Mat_Relese_Finish_Dt,PI_Mat_Relese_Deli_Status,PI_Addi_Article_No ,"
                sql = sql & "PI_Carrier,PI_Plan_Horiz_End,PI_Place_Consump,PI_Deli_Status_Number,PI_SegPadding,"

                sql = sql & "PKI_RecName,PKI_Version,PKI_Packing_No_Cust,PKI_Packing_No_Vend,PKI_Volume,PKI_SegPadding,"

                sql = sql & "DRT_RecName,DRT_Version ,DRT_Deli_Req_Text1,"

                If UBound(Data_Row) > 89 Then
                    sql = sql & "DRT_Deli_Req_Text2,DRT_Deli_Req_Text3,DRT_SegPadding,"
                    sql = sql & "SS_RecName ,SS_Version ,SS_counter_Segment511,"
                    sql = sql & "SS_counter_Segment512 ,SS_counter_Segment513,SS_counter_Segment514 ,SS_counter_Segment515,"
                    sql = sql & "SS_counter_Segment517 ,SS_counter_Segment518,SS_counter_Segment519 ,SS_SegPadding,"
                End If

                sql = sql & "slno,UNIT_CODE"

                sql = sql & " )"
                sql = sql & " Values"
                sql = sql & " ( "

                sql = sql & Val(trans_number) & ",302,"

                'GI
                sql = sql & Val(Data_Row(1)) & "," & Val(Data_Row(2)) & ",'" & Trim(Data_Row(3)) & "','" & Trim(Data_Row(4)) & "'," & Val(Data_Row(5)) & "," & Val(Data_Row(6)) & ",'" & FN_Date_Conversion(Trim(Data_Row(7))) & "',"
                sql = sql & Val(Data_Row(8)) & ",'" & Trim(Data_Row(9)) & "',"

                'GDRI
                sql = sql & Val(Data_Row(10)) & "," & Val(Data_Row(11)) & ",'" & Trim(Data_Row(12)) & "','" & Trim(Data_Row(13)) & "','" & FN_Date_Conversion(Trim(Data_Row(14))) & "','"
                sql = sql & Trim(Data_Row(15)) & "','" & FN_Date_Conversion(Trim(Data_Row(16))) & " ','" & Trim(Data_Row(17)) & "','" & Trim(Data_Row(18)) & "'," & Val(Data_Row(19)) & ",'"
                sql = sql & Trim(Data_Row(20)) & "','" & Trim(Data_Row(21)) & "','" & Trim(Data_Row(22)) & "','" & Trim(Data_Row(23)) & "'," & Val(Data_Row(24)) & "," & Val(Data_Row(25)) & ",'"
                sql = sql & Trim(Data_Row(26)) & "','" & Trim(Data_Row(27)) & "','" & Trim(Data_Row(28)) & "',"


                ''''''''''FOR DC
                If Data_Row(29) = "513" Then

                    sql = sql & "'',"

                    'DDRD

                    sql = sql & Val(Data_Row(29)) & "," & Val(Data_Row(30)) & ",'" & FN_Date_Conversion(Trim(Data_Row(31))) & "'," & Val(Data_Row(32)) & ",'" & FN_Date_Conversion(Trim(Data_Row(33))) & "',"
                    sql = sql & Val(Data_Row(34)) & "," & Val(Data_Row(35)) & ",'" & FN_Date_Conversion(Trim(Data_Row(36))) & "'," & Val(Data_Row(37)) & ",'" & FN_Date_Conversion(Trim(Data_Row(38))) & "'," & Val(Data_Row(39)) & ",'"
                    sql = sql & FN_Date_Conversion(Trim(Data_Row(40))) & "'," & Val(Data_Row(41)) & ",'" & FN_Date_Conversion(Trim(Data_Row(42))) & "'," & Val(Data_Row(43)) & ",'" & FN_Date_Conversion(Trim(Data_Row(44))) & "',"
                    sql = sql & Val(Data_Row(45)) & ",'" & Trim(Data_Row(46)) & "',"

                    'ADRD
                    sql = sql & Val(Data_Row(47)) & "," & Val(Data_Row(48)) & ",'" & FN_Date_Conversion(Trim(Data_Row(49))) & "'," & Val(Data_Row(50)) & ",'" & FN_Date_Conversion(Trim(Data_Row(51))) & "',"
                    sql = sql & Val(Data_Row(52)) & ",'" & FN_Date_Conversion(Trim(Data_Row(53))) & "'," & Val(Data_Row(54)) & ",'" & FN_Date_Conversion(Trim(Data_Row(55))) & "'," & Val(Data_Row(56)) & ",'"
                    sql = sql & FN_Date_Conversion(Trim(Data_Row(57))) & "'," & Val(Data_Row(58)) & ",'" & FN_Date_Conversion(Trim(Data_Row(59))) & "'," & Val(Data_Row(60)) & ",'" & FN_Date_Conversion(Trim(Data_Row(61))) & "',"
                    sql = sql & Val(Data_Row(62)) & ",'" & FN_Date_Conversion(Trim(Data_Row(63))) & "'," & Val(Data_Row(64)) & ",'" & Trim(Data_Row(65)) & "',"

                    'PI
                    sql = sql & Val(Data_Row(66)) & "," & Val(Data_Row(67)) & ",'" & FN_Date_Conversion(Trim(Data_Row(68))) & "','" & FN_Date_Conversion(Trim(Data_Row(69))) & "'," & Val(Data_Row(70)) & ",'"
                    sql = sql & FN_Date_Conversion(Trim(Data_Row(71))) & "','" & FN_Date_Conversion(Trim(Data_Row(72))) & "'," & Val(Data_Row(73)) & ",'" & Trim(Data_Row(74)) & "','" & Trim(Data_Row(75)) & "',"
                    sql = sql & Val(Data_Row(76)) & ",'" & Trim(Data_Row(77)) & "'," & Val(Data_Row(78)) & ",'" & Trim(Data_Row(79)) & "',"

                    'PKI
                    sql = sql & Val(Data_Row(80)) & "," & Val(Data_Row(81)) & ",'" & Trim(Data_Row(82)) & "','" & Trim(Data_Row(83)) & "'," & Val(Data_Row(84)) & ",'" & Trim(Data_Row(85)) & "',"

                    'DRT
                    sql = sql & Val(Data_Row(86)) & "," & Val(Data_Row(87)) & ",'" & Trim(Data_Row(88)) & "',"

                Else
                    ''''''''''FOR DC
                    sql = sql & "'" & Trim(Data_Row(29)) & "',"

                    'DDRD

                    sql = sql & Val(Data_Row(30)) & "," & Val(Data_Row(31)) & ",'" & FN_Date_Conversion(Trim(Data_Row(32))) & "'," & Val(Data_Row(33)) & ",'" & FN_Date_Conversion(Trim(Data_Row(34))) & "',"
                    sql = sql & Val(Data_Row(35)) & "," & Val(Data_Row(36)) & ",'" & FN_Date_Conversion(Trim(Data_Row(37))) & "'," & Val(Data_Row(38)) & ",'" & FN_Date_Conversion(Trim(Data_Row(39))) & "'," & Val(Data_Row(40)) & ",'"
                    sql = sql & FN_Date_Conversion(Trim(Data_Row(41))) & "'," & Val(Data_Row(42)) & ",'" & FN_Date_Conversion(Trim(Data_Row(43))) & "'," & Val(Data_Row(44)) & ",'" & FN_Date_Conversion(Trim(Data_Row(45))) & "',"
                    sql = sql & Val(Data_Row(46)) & ",'" & Trim(Data_Row(47)) & "',"

                    'ADRD
                    sql = sql & Val(Data_Row(48)) & "," & Val(Data_Row(49)) & ",'" & FN_Date_Conversion(Trim(Data_Row(50))) & "'," & Val(Data_Row(51)) & ",'" & FN_Date_Conversion(Trim(Data_Row(52))) & "',"
                    sql = sql & Val(Data_Row(53)) & ",'" & FN_Date_Conversion(Trim(Data_Row(54))) & "'," & Val(Data_Row(55)) & ",'" & FN_Date_Conversion(Trim(Data_Row(56))) & "'," & Val(Data_Row(57)) & ",'"
                    sql = sql & FN_Date_Conversion(Trim(Data_Row(58))) & "'," & Val(Data_Row(59)) & ",'" & FN_Date_Conversion(Trim(Data_Row(60))) & "'," & Val(Data_Row(61)) & ",'" & FN_Date_Conversion(Trim(Data_Row(62))) & "',"
                    sql = sql & Val(Data_Row(63)) & ",'" & FN_Date_Conversion(Trim(Data_Row(64))) & "'," & Val(Data_Row(65)) & ",'" & Trim(Data_Row(66)) & "',"

                    'PI
                    sql = sql & Val(Data_Row(67)) & "," & Val(Data_Row(68)) & ",'" & FN_Date_Conversion(Trim(Data_Row(69))) & "','" & FN_Date_Conversion(Trim(Data_Row(70))) & "'," & Val(Data_Row(71)) & ",'"
                    sql = sql & FN_Date_Conversion(Trim(Data_Row(72))) & "','" & FN_Date_Conversion(Trim(Data_Row(73))) & "'," & Val(Data_Row(74)) & ",'" & Trim(Data_Row(75)) & "','" & Trim(Data_Row(76)) & "',"
                    sql = sql & Val(Data_Row(77)) & ",'" & Trim(Data_Row(78)) & "'," & Val(Data_Row(79)) & ",'" & Trim(Data_Row(80)) & "',"

                    'PKI
                    sql = sql & Val(Data_Row(81)) & "," & Val(Data_Row(82)) & ",'" & Trim(Data_Row(83)) & "','" & Trim(Data_Row(84)) & "'," & Val(Data_Row(85)) & ",'" & Trim(Data_Row(86)) & "',"

                    'DRT
                    sql = sql & Val(Data_Row(87)) & "," & Val(Data_Row(88)) & ",'" & Trim(Data_Row(89)) & "',"
                End If

                If UBound(Data_Row) > 89 Then
                    'DRT
                    sql = sql & "'" & Trim(Data_Row(90)) & "','" & Trim(Data_Row(91)) & "','" & Trim(Data_Row(92)) & "',"

                    'SS
                    sql = sql & Val(Data_Row(93)) & "," & Val(Data_Row(94)) & "," & Val(Data_Row(95)) & "," & Val(Data_Row(96)) & ","
                    sql = sql & Val(Data_Row(97)) & "," & Val(Data_Row(98)) & "," & Val(Data_Row(99)) & "," & Val(Data_Row(100)) & ","
                    sql = sql & Val(Data_Row(101)) & "," & Val(Data_Row(102)) & ",'" & Trim(Data_Row(103)) & "',"
                End If
                sql = sql & Row & " , '" & gstrUNITID & "'"
                sql = sql & " ) "

                sqlCmd.CommandText = sql
                sqlCmd.ExecuteNonQuery()

                sqlitem = "select cust_drgno FROM CUSTITEM_MST " & _
                " WHERE CUST_DRGNO = '" & Trim(Data_Row(17)) & "'" & " AND active = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "' AND "

                If ShipmentFlag = True Then
                    sqlitem = sqlitem & " account_code = (SELECT" & " CUSTOMER_CODE FROM CUSTOMER_MST WHERE" & " CUST_VENDOR_CODE = '" & Trim(Data_Row(4)) & "' AND UNIT_CODE = '" & gstrUNITID & "' AND DOCK_CODE = '" & Trim(Data_Row(12)) & "' )"
                Else
                    sqlitem = sqlitem & " account_code = '" & Me.txtCustomerCode.Text & "'"
                End If

                sqlitem = sqlitem & " GROUP BY Cust_Drgno HAVING COUNT(CUST_DRGNO)  > 1"

                sqlCmd.CommandText = sqlitem
                sqlRDR = sqlCmd.ExecuteReader

                If sqlRDR.HasRows Then
                    While sqlRDR.Read
                        Msg = Msg & "'" + sqlRDR("CUST_DRGNO").ToString + "'" + vbCrLf
                    End While
                End If

                sqlRDR.Close()

                sql = "select Distinct dt,Cust_Code from Calendar_mkt_Cust " & " where dt IN ('" & FN_Date_Conversion(Trim(Data_Row(37))) & "', " & " '" & FN_Date_Conversion(Trim(Data_Row(39))) & "',  " & " '" & FN_Date_Conversion(Trim(Data_Row(41))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(43))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(45))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(50))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(52))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(54))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(56))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(58))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(60))) & " ', " & " '" & FN_Date_Conversion(Trim(Data_Row(62))) & " ', " & "  '" & FN_Date_Conversion(Trim(Data_Row(64))) & " ') " & " AND work_flg = 1 AND UNIT_CODE = '" & gstrUNITID & "' and Cust_Code = (SELECT CUSTOMER_CODE FROM " & " CUSTOMER_MST WHERE  UNIT_CODE = '" & gstrUNITID & "' AND CUST_VENDOR_CODE = '" & Trim(Data_Row(4)) & "' " & " AND DOCK_CODE = '" & Trim(Data_Row(12)) & "' ) "
                sqlCmd.CommandText = sql
                sqlRDR = sqlCmd.ExecuteReader

                If sqlRDR.HasRows Then
                    While sqlRDR.Read
                        If InStr(Replace(HOLIDAY, sqlRDR("cust_code").ToString + "---" + sqlRDR("dt").ToString, "$"), "$") = 0 Then
                            HOLIDAY = HOLIDAY & sqlRDR("cust_code").ToString + "---" + sqlRDR("dt").ToString & " " 'Replace used By Amit
                        End If
                    End While
                End If
                sqlRDR.Close()

                Row = Row + 1
                range = Obj_EX.Cells(Row, 1)
                If Not range.Value Is Nothing Then
                    Cell_Data = (range.Value.ToString)
                Else
                    Cell_Data = ""
                End If

            End While

        End If

        Dim countREC As Short

        sql = "select DISTINCT C.ACCOUNT_CODE, C.cust_drgno,COUNT(C.item_code) countitem from custitem_mst C with (nolock) where C.active = 1 AND SCHUPLDREQD = 1 AND C.UNIT_CODE = '" & gstrUNITID & "' and "
        If ShipmentFlag = True Then
            sql = sql & " C.ACCOUNT_CODE IN (SELECT DISTINCT CONSIGNEE_CODE"
            If Darwin_FileType = "VDA" Then
                sql = sql & " FROM VW_SCHEDULE_PROPOSAL with (nolock) WHERE UNIT_CODE = '" & gstrUNITID & "') AND C.CUST_DRGNO IN (SELECT DISTINCT CUST_DRGNO"
            End If
        Else
            sql = sql & " C.ACCOUNT_CODE = '" & txtCustomerCode.Text & "'"
            sql = sql & " AND C.CUST_DRGNO IN (SELECT DISTINCT CUST_DRGNO"
        End If

        sql = sql & " FROM VW_SCHEDULE_PROPOSAL with (nolock)"
        sql = sql & " WHERE DOC_NO = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "') " & " group by C.Account_code, C.cust_drgno"

        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader
        Msg = ""

        If sqlRDR.HasRows Then
            While sqlRDR.Read
                countREC = sqlRDR("COUNTITEM").ToString
                If countREC > 1 Then
                    Msg = Msg & "  " + sqlRDR("cust_drgno").ToString
                    Flag = 1
                End If
            End While

            If Msg <> "" Then
                MsgBox("Following Cust_DrgNo(s) Are Active For Multiple Items " & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
        End If
        sqlRDR.Close()

        sql = "select ltrim(rtrim(item_code)) + '    ' + ltrim(rtrim(cust_drgno)) as item_custdrgno From custitem_mst Where Active = 1 AND UNIT_CODE = '" & gstrUNITID & "' AND SCHUPLDREQD = 1 " & " and ltrim(rtrim(item_code)) + '    ' + ltrim(rtrim(cust_drgno))" & " in (select distinct ltrim(rtrim(item_code)) + '    ' + ltrim(rtrim(cust_drgno))" & " from vw_schedule_proposal with (nolock)" & " where doc_no = " & trans_number & " AND UNIT_CODE = '" & gstrUNITID & "') and ltrim(rtrim(item_code)) + '    ' + ltrim(rtrim(cust_drgno))" & " not in (select ltrim(rtrim(item_code)) + '    ' + ltrim(rtrim(cust_drgno))" & " as item_custdrgno From custitem_mst" & " where account_code = '" & txtCustomerCode.Text & "' and active = 1 AND SCHUPLDREQD = 1 AND UNIT_CODE = '" & gstrUNITID & "')"


        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader

        Msg = ""
        If sqlRDR.HasRows Then
            While sqlRDR.Read
                Msg = Msg & "'" + sqlRDR("item_custdrgno").ToString + "'" + vbCrLf
            End While

            If Len(Trim(Msg)) > 0 Then
                MsgBox("Following Items Are Not Defined In The System" & vbCrLf & "for Customer " & txtCustomerCode.Text & " : " & vbCrLf & Msg, MsgBoxStyle.OkOnly, ResolveResString(100))
                Flag = 1
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            End If

            Trans_Satus = False

        End If
        sqlRDR.Close()

        HOLIDAY = Replace(HOLIDAY, " ", Chr(13))
        If HOLIDAY <> "" Then
            MsgBox("Following is/are not working day(s):" & vbCrLf & vbCrLf & "Consignee---Date" & vbCrLf & HOLIDAY & vbCrLf)
            Flag = 1
        End If
        If ShipmentFlag = True Then
            sql = "select DISTINCT D.GI_Vend_Code as GI_Vend_Code from schedule_upload_darwin_DTL_temp D," & "scheduleparameter_mst s ,schedule_upload_darwin_HDR_temp H " & " where D.GI_Vend_Code  not in(select S.wh_code  from scheduleparameter_mst s " & " where s.customer_code =  '" & txtCustomerCode.Text & "' AND S.UNIT_CODE = '" & gstrUNITID & "') and " & " H.Cust_Code = '" & txtCustomerCode.Text & "' " & " AND H.DOC_NO = D.DOC_NO AND H.UNIT_CODE = D.UNIT_CODE   AND s.Customer_code = h.cust_code AND H.UNIT_CODE = s.UNIT_CODE  " & " and D.doc_no = " & trans_number & " AND D.UNIT_CODE = '" & gstrUNITID & "' "
            sqlCmd.CommandText = sql
            sqlRDR = sqlCmd.ExecuteReader

            If sqlRDR.HasRows Then
                sqlRDR.Read()
                msgWH = msgWH & "  '" + sqlRDR("GI_Vend_Code").ToString + "'  "
                Flag = 1
            End If
            sqlRDR.Close()

            If msgWH <> "" Then
                MsgBox("WRONG WAREHOUSE: " & msgWH, MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
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

        If Flag = 1 Then
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Exit Sub
        Else
            If ShipmentFlag = False And chkDlyPullQty.Checked = True Then
                MsgBox("Shipment For This Customer Is Not Through Warehouse, " & vbCrLf & "So, You Can't Use Daily Pull Qty For Safety Stock Calculations.", vbInformation, ResolveResString(100))
                Obj_FSO = Nothing
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                Exit Sub
            Else
                Me.txtDocNo.Text = trans_number
            End If
        End If

        If Flag = 0 Then
            sqlTrans = sqlCmd.Connection.BeginTransaction()
            sqlCmd.Transaction = sqlTrans
            isTrans = True

            sql = "INSERT INTO schedule_upload_DARWIN_hdr SELECT * FROM schedule_upload_DARWIN_hdr_TEMP where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "INSERT INTO schedule_upload_DARWIN_DTL SELECT * FROM schedule_upload_DARWIN_DTL_TEMP where doc_no = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sql = "update documenttype_mst set current_no = " & trans_number & " where Doc_Type = 302 and GETDATE() between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandText = sql
            sqlCmd.ExecuteNonQuery()

            sqlTrans.Commit()
            isTrans = False
        End If

        If ShipmentFlag = True Then
            sql = "EXEC SP_CALCULATESAFETYSTOCKFORSCHEDULE '" & gstrUNITID & "', '" & Me.txtCustomerCode.Text & "'" & ",'" & trans_number & "'," & "'" & WA & "','" & sch & "','" & gstrIpaddressWinSck & "'"
            sqlCmd.CommandText = sql
            sqlCmd.CommandTimeout = 0
            sqlCmd.ExecuteNonQuery()

            sqlCmd.CommandText = "Delete from tmp_Schedule_Uploading_Darwin WHERE UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.ExecuteNonQuery()

            sqlCmd.CommandText = "insert into tmp_Schedule_Uploading_Darwin select * from vw_Schedule_Uploading_Darwin where Doc_NO = '" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.ExecuteNonQuery()

            RSWHCOUNT = New SqlCommand
            RSWHCOUNT.Connection = SqlConnectionclass.GetConnection

            RSWHCOUNTSTOCK = New SqlCommand
            RSWHCOUNTSTOCK.Connection = SqlConnectionclass.GetConnection

            sql = "Select count(distinct GI_Vend_Code) COUNT,GI_Vend_Code as WH_CODE from tmp_Schedule_Uploading_Darwin where doc_no='" & trans_number & "' AND UNIT_CODE = '" & gstrUNITID & "' group by GI_Vend_Code"
            RSWHCOUNT.CommandText = sql
            RDRWHCOUNT = RSWHCOUNT.ExecuteReader

            sql = "Select count(distinct GI_Vend_Code) COUNT,GI_Vend_Code as WH_CODE from tmp_Schedule_Uploading_Darwin  where  UNIT_CODE = '" & gstrUNITID & "' AND GI_Vend_Code not in (Select distinct WareHouse_Code from WareHouse_Stock_Dtl WHERE UNIT_CODE = '" & gstrUNITID & "' ) group by GI_Vend_Code"
            RSWHCOUNTSTOCK.CommandText = sql
            RDRWHCOUNTSTOCK = RSWHCOUNTSTOCK.ExecuteReader

            If RDRWHCOUNT.HasRows And RDRWHCOUNTSTOCK.HasRows Then
                If RDRWHCOUNT("COUNT").ToString > 0 And RDRWHCOUNT("COUNT").ToString = RDRWHCOUNTSTOCK("COUNT").ToString And RDRWHCOUNT("WH_CODE").ToString = RDRWHCOUNTSTOCK("WH_CODE").ToString Then
                    MsgBox("Stock is not defined for the Warehouse(s).So no Schedule will be proposed.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If
                    Exit Sub
                End If
                If RDRWHCOUNT("COUNT").ToString > 1 Then
                    MsgBox("You have the Release File with more than 1 warehouses." & vbCrLf & "Details for these will be available in Schedule Proposal Details.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            End If

            RDRWHCOUNT.Close()
            RDRWHCOUNTSTOCK.Close()

            RSWHCOUNT.Dispose()
            RSWHCOUNT = Nothing

            RSWHCOUNTSTOCK.Dispose()
            RSWHCOUNTSTOCK = Nothing

        End If

        If ShipmentFlag = True Then
            If chkdaywisesch.Checked = True Then
                Call FN_TRANSFERDATAINCOVISINT(trans_number, Darwin_FileType)
                Darwin_FileType = "COVISINT"
            End If

            Call FN_Display(trans_number, Darwin_FileType)

        Else
            If chkdaywisesch.Checked = True Then
                Call FN_TRANSFERDATAINCOVISINT(trans_number, Darwin_FileType)
                Darwin_FileType = "COVISINT"
            End If


            sqlTrans = sqlCmd.Connection.BeginTransaction
            sqlCmd.Transaction = sqlTrans
            isTrans = True

            Call FN_Display_WITHOUTWH(trans_number)

            sqlTrans.Commit()
            isTrans = False

        End If

        sqlCmd.CommandText = "set dateformat 'dmy'"
        sqlCmd.ExecuteNonQuery()


        sql = "Select Top 1 Doc_No From ScheduleProposalcalculations Where Doc_No= '" & Val(trans_number) & "' and shipment_qty > 0 AND UNIT_CODE = '" & gstrUNITID & "'"
        sqlCmd.CommandText = sql
        sqlRDR = sqlCmd.ExecuteReader

        If Not sqlRDR.HasRows Then
            sqlRDR.Close()
            YesNo = CStr(MsgBox("No Schedule Proposed." + vbCrLf + "Do You Want To Move Files To BackUp Folder?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)))

            If YesNo = CStr(MsgBoxResult.Yes) Then
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                lblMessage1.Text = "No Schedule Proposed."
                Call MoveFile()
            End If
        Else
            sqlRDR.Close()
            lblMessage1.Text = "Schedule has been Uploaded Succesfully."

            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            sqlCmd.Dispose()
            sqlCmd = Nothing

            Call Updt_DailyMkt(Darwin_FileType)

            If mblnfilemove = False Then
                Call MoveFile()
            End If

        End If
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        TmpRs = Nothing
        Exit Sub
ERR_Renamed:

        If isTrans = True Then
            sqlTrans.Rollback()
            sqlTrans = Nothing
        End If

        If Not sqlCmd Is Nothing Then
            sqlCmd.Dispose()
            sqlCmd = Nothing
        End If

        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub ctlFormHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        ShowHelp(("underconstruction.htm"))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub DTPicker1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTPicker1.Validating
        On Error GoTo ErrHandler

        Dim rsWHDate As New ClsResultSetDB

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

        rsWHDate.GetResult("SELECT MAX(TRANS_DT) as TRANS_DT FROM WAREHOUSE_STOCK_DTL" & " WHERE WAREHOUSE_CODE = '" & txtUnitCode.Text & "' " & " AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'")

        If VB6.Format(DTPicker1.Value, "YYYYMMMDD") < VB6.Format(rsWHDate.GetValue("TRANS_DT"), "YYYYMMMDD") Then
            MsgBox("You Have Already Uploaded Stock For" + vbCrLf + rsWHDate.GetValue("trans_dt"), vbInformation + vbOKOnly, ResolveResString(100))
            DTPicker1.Format = DateTimePickerFormat.Custom
            DTPicker1.CustomFormat = gstrDateFormat
            DTPicker1.Value = GetServerDate()
        End If

        If Me.DTPicker1.Value > GetServerDate() Then
            DTPicker1.Format = DateTimePickerFormat.Custom
            DTPicker1.CustomFormat = gstrDateFormat
            Me.DTPicker1.Value = GetServerDate()
        End If

        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)


    End Sub

    Private Sub FRMMKTTRN0028_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Me.Tag) = True
        DTPicker1.Format = DateTimePickerFormat.Custom
        DTPicker1.CustomFormat = gstrDateFormat
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

            GroupBox1.Visible = True
            ChkDaimler.Checked = True
            ChkFord.Checked = False
            ChkDaimler.Enabled = False
            ChkFord.Enabled = False
        End If

        If Me.OptReleaseFile.Checked = True Then
            Me.lblTransitDays.Visible = False
            Me.lblTransitDays.Text = "Transit Days By Sea"
            Me.lbltransitdaysvalue.Visible = False
        Else
            Me.lblTransitDays.Visible = False
            Me.lbltransitdaysvalue.Visible = False
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub FRMMKTTRN0028_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FRMMKTTRN0028_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            ctlFormHeader_ClickEvent(ctlFormHeader, New System.EventArgs())
        End If

    End Sub

    Private Sub FRMMKTTRN0028_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub FRMMKTTRN0028_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        bln_dateCheck = True
        Call FitToClient(Me, frmMain, ctlFormHeader, frmButton, 450)
        Call FillLabelFromResFile(Me)
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)

        Call EnableControls(True, Me)
        Me.KeyPreview = True
        Me.lblDocno.Visible = False
        Me.txtDocNo.Visible = False
        Me.OptWareHouseFile.Checked = True
        Me.OptStock.Checked = True

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FRMMKTTRN0028_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler

        Me.Dispose()
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub optAvgofNextMonths_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAvgofNextMonths.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ERR_Renamed

            txtNoOfMonths.Text = ""
            txtNoOfMonths.Enabled = True

            Exit Sub
ERR_Renamed:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Sub

        End If
    End Sub

    Private Sub optCurMonthSch_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCurMonthSch.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ERR_Renamed

            txtNoOfMonths.Text = ""
            txtNoOfMonths.Enabled = False

            Exit Sub
ERR_Renamed:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Sub


        End If
    End Sub

    Private Sub optNextMonthSch_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNextMonthSch.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ERR_Renamed

            txtNoOfMonths.Text = ""
            txtNoOfMonths.Enabled = False

            Exit Sub
ERR_Renamed:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Sub

        End If
    End Sub

    Private Sub OptReleaseFile_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptReleaseFile.CheckedChanged
        If eventSender.Checked Then

            On Error GoTo ERR_Renamed
            Call CmdClear_Click(CmdClear, New System.EventArgs())
            chkDlyPullQty.Visible = True
            lblUnitCode.Text = "Plant Code"
            lbldate.Enabled = False
            DTPicker1.Enabled = False
            Me.lbltransitdaysvalue.Visible = False
            Me.lblTransitDays.Visible = False
            Me.lblTransitDays.Text = "Transit Days By Sea"
            Me.Frame3.Enabled = True
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
            'Call AlignGrID()
            CmdUploadCSV.Enabled = False
            Me.lblDocno.Visible = True
            Me.txtDocNo.Visible = True
            Me.frmFileoption.Visible = False

            Me.chkdaywisesch.Visible = True

            GroupBox1.Visible = False

            Exit Sub
ERR_Renamed:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Sub

        End If
    End Sub

    Private Sub OptWareHouseFile_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptWareHouseFile.CheckedChanged
        If eventSender.Checked Then

            On Error GoTo ERR_Renamed

            txtCustomerCode.Text = "" : txtUnitCode.Text = "" : txtFileName.Text = ""

            lblUnitCode.Text = "Ware House Code"
            lbldate.Enabled = True
            DTPicker1.Enabled = True

            chkDlyPullQty.Visible = False
            Me.lbltransitdaysvalue.Visible = False
            Me.lblTransitDays.Visible = False

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

            Me.chkdaywisesch.Visible = False

            GroupBox1.Visible = True

            Exit Sub
ERR_Renamed:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Sub

        End If
    End Sub

    Private Sub txtConsignee_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtConsignee.GotFocus
        consFocus = False
    End Sub

    Private Sub txtConsignee_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConsignee.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdConsigneeHelp_Click(cmdConsigneeHelp, New System.EventArgs())
        End If
    End Sub

    Private Sub txtConsignee_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtConsignee.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case Keys.Enter
                bool_Validate_Cons = False
                txtConsignee_Validating(txtConsignee, New System.ComponentModel.CancelEventArgs((False)))
        End Select
    End Sub

    Private Sub txtConsignee_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtConsignee.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo Errorhandler

        consFocus = False

        If bool_Validate_Cons = True Then
            bool_Validate_Cust = False
            Exit Sub
        End If

        bool_Validate_Cust = True
        Dim rsobject As New ClsResultSetDB
        If txtCustomerCode.Text = "" Then
            MsgBox("Please Enter Customer Code", MsgBoxStyle.OkOnly, ResolveResString(100))
            Call txtCustomerCode_TextChanged(txtCustomerCode, New System.EventArgs())
            txtCustomerCode.Focus()
            GoTo EventExitSub
        End If
        If Len(Trim(Me.txtConsignee.Text)) > 0 Then
            If Not CheckRecord("Select Customer_code from Customer_mst where Customer_code = '" & Me.txtConsignee.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'") Then
                MsgBox(" Invalid Consignee Code", MsgBoxStyle.Information, ResolveResString(100))
                Me.txtConsignee.Text = "" : Me.txtConsignee.Focus()
                Cancel = True
            End If
        End If

        If Cancel = True Then
            Me.txtConsignee.Focus()
        Else
            Call rsobject.GetResult("Select Customer_Mst.Cust_Name From Customer_Mst  Where Customer_Mst.Customer_Code= '" & Trim(Me.txtConsignee.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "'")
            Me.txtUnitCode.Focus()
        End If

        rsobject.ResultSetClose()
        rsobject = Nothing

        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)


EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        custFocus = False
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged

        LblCustomerName.Text = ""
        lbltransitdaysvalue.Text = ""
        Me.txtUnitCode.Text = ""
        txtConsignee.Text = ""
        Me.DTPicker1.Value = GetServerDate()
        Me.lblUnitName.Text = ""
        Me.txtFileName.Text = ""
        bool_Validate_Cust = False
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdCustHelp_Click(cmdCustHelp, New System.EventArgs())
        End If
    End Sub

    Private Sub TxtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            bool_Validate_Cust = False
            TxtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs((False)))

        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo Errorhandler

        custFocus = False
        If bool_Validate_Cust = True Then
            bool_Validate_Cust = False
            Exit Sub
        End If

        Dim rsobject As New ClsResultSetDB
        mblnDailymktUpdated = False
        mblnfilemove = False
        bool_Validate_Cust = True

        If Len(Trim(Me.txtCustomerCode.Text)) > 0 Then
            If Not CheckRecord("Select Customer_code from ScheduleParameter_mst where Customer_code = '" & Me.txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'") Then
                MsgBox(" Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                Me.LblCustomerName.Text = "" : Me.lbltransitdaysvalue.Text = ""
                Me.txtCustomerCode.Text = "" : Me.txtCustomerCode.Focus()
                Cancel = True
            End If
        End If

        If Cancel = True Then
            Me.txtCustomerCode.Focus()
        Else
            Call rsobject.GetResult("Select Customer_Mst.Cust_Name,ScheduleParameter_mst.TransitDaysbysea From ScheduleParameter_mst,Customer_Mst  Where Customer_Mst.Customer_Code=ScheduleParameter_mst.Customer_Code And Customer_Mst.Customer_Code = '" & Trim(Me.txtCustomerCode.Text) & "' AND Customer_Mst.UNIT_CODE=ScheduleParameter_mst.UNIT_CODE AND Customer_Mst.UNIT_CODE = '" & gstrUNITID & "'")
            If rsobject.GetNoRows > 0 Then
                Me.LblCustomerName.Text = rsobject.GetValue("Cust_Name")
                Me.lbltransitdaysvalue.Text = rsobject.GetValue("TransitDaysBySea")
            End If
            Me.txtConsignee.Focus()

            If OptReleaseFile.Checked = True Then
                txtCustomerCode.Enabled = False
                cmdCustHelp.Enabled = False
                Call CmdUploadCSV_Click(CmdUploadCSV, New System.EventArgs())
                CmdClear.Focus()
            End If

        End If

        rsobject.ResultSetClose()
        rsobject = Nothing

        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Function CheckRecord(ByRef strsql As String) As Boolean
        On Error GoTo Errorhandler

        Dim rsobject As New ClsResultSetDB

        rsobject.GetResult(strsql)
        If rsobject.RowCount > 0 Then
            CheckRecord = True
        Else
            CheckRecord = False
        End If
        rsobject.ResultSetClose()
        rsobject = Nothing

        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function

    Private Sub txtFileName_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFileName.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdFileHelp_Click(cmdFileHelp, New System.EventArgs())
        End If
    End Sub

    Private Sub txtNoOfMonths_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNoOfMonths.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Errorhandler

        If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not KeyAscii = 8 Then
            KeyAscii = 0
        End If

        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtUnitCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitCode.Enter
        unitFocus = False
    End Sub

    Private Sub txtUnitCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitCode.GotFocus
        unitFocus = False
    End Sub

    Private Sub TxtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged

        lblUnitName.Text = ""

    End Sub

    Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdUnitHelp_Click(cmdUnitHelp, New System.EventArgs())
        End If
    End Sub

    Private Sub TxtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Call TxtUnitCode_Validating(txtUnitCode, New System.ComponentModel.CancelEventArgs(False))
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo Errorhandler
        Dim rsobject As New ClsResultSetDB
        unitFocus = False
        If OptReleaseFile.Checked Then
            If Len(Trim(Me.txtUnitCode.Text)) > 0 Then
                If Not CheckRecord("Select plant_c from plant_mst where plant_c = '" & Me.txtUnitCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'") Then
                    MsgBox(" Invalid Unit Code", MsgBoxStyle.Information, ResolveResString(100))
                    Me.txtUnitCode.Text = "" : Me.txtUnitCode.Focus()
                    Cancel = True
                End If
            End If

            If Cancel = True Then
                Me.txtUnitCode.Focus()
            Else
                Call rsobject.GetResult("Select plant_nm from plant_mst where plant_c = '" & Me.txtUnitCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'")
                lblUnitName.Text = rsobject.GetValue("plant_nm")
                Me.txtFileName.Focus()
            End If
        End If
        If OptWareHouseFile.Checked And txtUnitCode.Text <> "" Then
            rsobject.GetResult("select C.wh_code,W.WH_DESCRIPTION  from custwarehouse_mst C,WAREHOUSE_MST W " & " where C.customer_code = '" & txtConsignee.Text & "' AND C.WH_CODE = W.WH_CODE " & " and active = 1 AND C.UNIT_CODE = W.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'")

            If rsobject.GetNoRows = 0 Then
                MsgBox("Invalid Warehouse Code", MsgBoxStyle.OkOnly, ResolveResString(100))
                txtUnitCode.Text = ""
            End If
            If rsobject.GetNoRows > 0 Then
                rsobject.GetResult("select top 1 WarehouseFile_Location" & " from scheduleparameter_mst" & " where customer_code = '" & txtCustomerCode.Text & "' and WH_Code='" & Trim(txtUnitCode.Text) & "' and consignee_code='" & Trim(txtConsignee.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "' order by entdt")
                txtFileName.Text = rsobject.GetValue("WarehouseFile_Location")
                txtFileName.Focus()
            End If
        End If
        rsobject.ResultSetClose()
        rsobject = Nothing
        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub WareHouse_Inv_Upload()
        Dim invobj As New prj_uploadInvoiceDaimler.prj_uploadInvoiceDaimler

        Dim RowNo As Short
        Dim RSrow As New ClsResultSetDB

        On Error GoTo ERR_Renamed
        Dim inv_no As Object = Nothing, ExpInv_No As Object = Nothing
        Dim sql As String = ""
        Dim Col, Row As Short
        Dim Stk_Qty, Item_Rate As Double
        Dim rsRevno As ClsResultSetDB
        Dim lngRevno As Integer = 0
        Dim rsValidateInvoiceno As ClsResultSetDB
        Dim Msg As String = ""
        Dim InvDt As String = ""

        If UCase(Mid(Trim(txtFileName.Text), Trim(txtFileName.Text).Length - 2, 3)) = UCase("txt") Then
            Msg = invobj.WareHouse_Inv_TextFileUpload(txtCustomerCode.Text, txtUnitCode.Text, txtConsignee.Text, DTPicker1.Value, gstrConnectSQLClient, Trim(txtFileName.Text)).ToString
            MsgBox(Msg)
            Exit Sub
        End If

        sql = "select StartingRow from scheduleparameter_mst where customer_code = '" & txtCustomerCode.Text & "' and wh_code = '" & txtUnitCode.Text & "' and Consignee_code='" & Trim(txtConsignee.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "'"
        RSrow.GetResult(sql)

        If RSrow.GetNoRows > 0 Then
            If RSrow.GetValue("StartingRow") = "" Or RSrow.GetValue("StartingRow") Is System.DBNull.Value Then
                RowNo = 1
            Else
                RowNo = RSrow.GetValue("StartingRow")
            End If
        End If
        RSrow.ResultSetClose()

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
        rsRevno = New ClsResultSetDB
        rsRevno.GetResult("Select max(revno) as Maxrevno from InvoiceUpldWH where customer_code='" & Me.txtCustomerCode.Text & "'" & " and WareHouseCode='" & Me.txtUnitCode.Text & "' AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'")

        If rsRevno.GetNoRows > 0 Then
            lngRevno = IIf(rsRevno.GetValue("Maxrevno") = "", 0, rsRevno.GetValue("Maxrevno"))
        Else
            lngRevno = 0
        End If
        rsRevno.ResultSetClose()
        rsRevno = Nothing
        lngRevno = lngRevno + 1
        mP_Connection.BeginTrans()
        mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        While Len(ExpInv_No) <> 0

            If Len(inv_no) <> 0 Then
                mP_Connection.Execute("insert into InvoiceUpldWH(" & "customer_code,WareHouseCode,Inv_dt,Invoice_no,revno,upld_dt,ent_dt,CONSIGNEE_CODE,UNIT_CODE)" & "values('" & Me.txtCustomerCode.Text & "'," & " '" & Me.txtUnitCode.Text & "' , " & " '" & InvDt & "' , " & " '" & inv_no & " '," & lngRevno & ",'" & DTPicker1.Value & "',getdate(),'" & txtConsignee.Text & "','" & gstrUNITID & "') ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
        rsValidateInvoiceno = New ClsResultSetDB

        rsValidateInvoiceno.GetResult("select Invoice_no from InvoiceUpldWH " & " where REVNO= " & lngRevno & " AND Invoice_no not in (" & " select CAST (doc_no AS VARCHAR) from saleschallan_dtl WHERE UNIT_CODE = '" & gstrUNITID & "')" & " and warehousecode = '" & txtUnitCode.Text & "'" & " and customer_code = '" & txtCustomerCode.Text & "'" & " AND INV_DT >= '01 Jan 2008' AND CONSIGNEE_CODE = '" & txtConsignee.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'") 'Inv_Dt Condition Added On 08 Apr 2008
        If rsValidateInvoiceno.GetNoRows > 0 Then
            rsValidateInvoiceno.MoveFirst()
            While Not rsValidateInvoiceno.EOFRecord
                Msg = Msg + rsValidateInvoiceno.GetValue("Invoice_no") + " ,"
                rsValidateInvoiceno.MoveNext()
            End While
            rsValidateInvoiceno.ResultSetClose()
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
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

        Exit Sub
ERR_Renamed:

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing

        End If

        Exit Sub

    End Sub

    Private Sub cmdConsigneeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConsigneeHelp.Click

        On Error GoTo ErrHandler

        Dim strDocNoHelp() As String = Nothing
        If OptWareHouseFile.Checked = True Then
            If txtCustomerCode.Text = "" Then
                MsgBox("Please Enter Customer Code.", MsgBoxStyle.OkOnly, ResolveResString(100))
                txtConsignee.Text = ""
                txtCustomerCode.Focus()
                txtCustomerCode_TextChanged(txtCustomerCode, New System.EventArgs())
                Exit Sub
            End If
            strDocNoHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code, c.cust_name from customer_mst c WHERE C.UNIT_CODE = '" & gstrUNITID & "' ", "List of Customers")

        ElseIf OptReleaseFile.Checked = True Then
            strDocNoHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct c.customer_code,c.cust_name from customer_mst c WHERE C.UNIT_CODE = '" & gstrUNITID & "'", " List of Customer ")
        End If


        If UBound(strDocNoHelp) <> -1 Then
            If strDocNoHelp(0) <> "0" Then
                Me.txtConsignee.Text = strDocNoHelp(0)
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub upload_covisint(ByRef trans_number As String, ByRef SheetCount As Short, ByVal CONSIGNEE_CODE As String)
        On Error GoTo ErrHandler

        Dim Flag As Short

        Dim Cell_Data As String
        Dim Row, i As Short
        Dim Data_Row() As String
        Dim Trans_Satus As Boolean
        Dim Upload_FileType, sql, Cell_Data1 As String
        Dim Rev_No, Col As Short

        Dim rsItems As New ClsResultSetDB
        Dim RSdt As New ClsResultSetDB
        Dim HOLIDAY As String
        Dim Item, Msg As String
        Dim sql1, sql2 As String
        Dim rsObjectWH As New ClsResultSetDB

        Dim sqlConsignee As String
        Dim RSconsignee As New ClsResultSetDB
        Dim whCode As String
        Dim FactoryCode As String

        Dim sqlWarehouse As String
        Dim msgWH As String


        HOLIDAY = ""
        msgWH = ""

        If Obj_FSO Is Nothing Then
            Obj_FSO = New Scripting.FileSystemObject
        End If

        If Obj_FSO.FileExists(Me.txtFileName.Text) = False Then
            MsgBox(" File Does not Exist ", MsgBoxStyle.Information, ResolveResString(100))
            txtFileName.Focus()
            Exit Sub
        End If

        Obj_EX.Sheets.Item(SheetCount).Select()

        Row = 1
        range = Obj_EX.Cells(Row, 1)
        If Not range.Value Is Nothing Then
            Cell_Data = (range.Value.ToString)
        Else
            Cell_Data = ""
        End If


        If Len(Cell_Data) = 0 Then
            Obj_FSO = Nothing
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        Row = 6
        i = 2

        If ShipmentFlag = False Then
            whCode = ""
            FactoryCode = ""
        Else
            range = Obj_EX.Cells(2, 2)
            whCode = range.Value.ToString
            range = Obj_EX.Cells(3, 2)
            FactoryCode = range.Value.ToString
        End If

        Row = 6
        Col = 2


        While Not Obj_EX.Cells(5, Col).VALUE Is Nothing
            While Not Obj_EX.Cells(Row, 1).VALUE Is Nothing
                sql = "Insert into schedule_upload_covisint_dtl_temp(doc_no,doc_type, " & _
                    "item_code,WH_CODE,factory_code,consignee_code,delivery_date,qty,ent_dt,ent_uid,updt_dt,updt_uid,UNIT_CODE)" & _
                    "Values (" & trans_number & ",302,'" & Obj_EX.Cells(5, Col).value & "'," & _
                    " '" & whCode & "','" & FactoryCode & "','" & CONSIGNEE_CODE & "','" & VB6.Format(Obj_EX.Cells(Row, 1).Value, "dd MMM yyyy") & "'" & _
                    " ,'" & Obj_EX.Cells(Row, Col).value & "',getDate()," & _
                    " '" & mP_User & "',getDate() ,'" & mP_User & "','" & gstrUNITID & "')"
                mP_Connection.Execute(sql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Row = Row + 1
            End While
            If Col = Obj_EX.Columns.Count Then
                Exit While
            Else
                Col = Col + 1
            End If

            Row = 6


        End While


        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Obj_FSO = Nothing
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

    End Sub

    Private Function FN_Display_WITHOUTWH(ByVal TRANS_NUMBER As String) As Object
        On Error GoTo Err

        Dim Transit_Days, Row, Buffer_Days As Integer
        Dim adors As New ADODB.Recordset
        Dim CustDrgNo As Object = Nothing, DELDT As Object = Nothing
        Dim SCHQTY As Double
        Dim oCmd As ADODB.Command
        Dim SFTYDAYS_MNTD As Object = Nothing
        Dim SAFETYDAYS_BELOW As Long
        Dim TMPWHSTOCK As Long
        Dim sql As String = "", updSQL As String = "", strWHCode As String = ""
        Dim RSWHSTOCK As New ClsResultSetDB
        Dim rsDate As ADODB.Recordset
        Dim rsbagqty As New ClsResultSetDB
        Dim Rs As ADODB.Recordset
        Dim rsTransitDays As ADODB.Recordset

        If Darwin_FileType = "EDIFACT" Or Darwin_FileType = "COVISINT" Then

            If Darwin_FileType = "EDIFACT" Then
                sql = "Select Distinct D.Delivery_Dt,h.PARTY_ID1 ,H.PARTY_ID3," & _
                      " C.Cust_Drgno,I.Item_Code,I.DESCRIPTION, " & _
                       "QUANTITY AS SHIPQTY,Frequency,DISPATCH_PATTERN"
                sql = sql & " From SCHEDULE_UPLOAD_DARWINEDIFACT_DTL D, CUSTITEM_MST C,ITEM_MST I,  " & _
                    "SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H, SCHEDULEPARAMETER_mst SP"
                sql = sql & " Where C.Account_code=SP.Customer_code AND C.UNIT_CODE = SP.UNIT_CODE and C.Active=1 AND C.SCHUPLDREQD = 1"
                sql = sql & " And H.Doc_No=" & TRANS_NUMBER & ""
                sql = sql & " AND D.ITEM_CODE = C.CUST_DRGNO AND D.UNIT_CODE = C.UNIT_CODE AND C.ITEM_CODE = I.ITEM_CODE AND C.UNIT_CODE = I.UNIT_CODE"
                sql = sql & " AND D.DOC_NO = H.DOC_NO AND D.UNIT_CODE = H.UNIT_CODE and ltrim(rtrim(frequency))<>'' and ltrim(rtrim(dispatch_pattern))<>'' " & _
                            " AND SP.CUSTOMER_CODE  = '" & Me.txtCustomerCode.Text & "' AND D.UNIT_CODE = '" & gstrUNITID & "'"
                sql = sql & " Order By D.DELIVERY_DT "
            End If

            If Darwin_FileType = "COVISINT" Then
                sql = "Select Distinct D.Delivery_DATE AS DELIVERY_DT,D.WH_CODE AS PARTY_ID1,C.Cust_Drgno,I.Item_Code," & _
                  " I.DESCRIPTION, QTY AS SHIPQTY,FACTORY_CODE " & _
                  " From SCHEDULE_UPLOAD_COVISINT_DTL D, CUSTITEM_MST C,ITEM_MST I," & _
                  " SCHEDULE_UPLOAD_COVISINT_HDR H, SCHEDULEPARAMETER_mst SP" & _
                  " Where C.Account_code=SP.Customer_code AND C.UNIT_CODE = SP.UNIT_CODE and" & _
                  " C.Active=1 AND C.SCHUPLDREQD = 1 And H.Doc_No = " & TRANS_NUMBER & "  AND D.ITEM_CODE = C.CUST_DRGNO AND D.UNIT_CODE = C.UNIT_CODE" & _
                  " AND C.ITEM_CODE = I.ITEM_CODE AND D.DOC_NO = H.DOC_NO AND C.UNIT_CODE = I.UNIT_CODE AND D.UNIT_CODE = H.UNIT_CODE" & _
                  " and SP.CUSTOMER_CODE  = '" & txtCustomerCode.Text & "' AND D.UNIT_CODE = '" & gstrUNITID & "'  Order By D.DELIVERY_DATE"
            End If

            adors = New ADODB.Recordset
            adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            SCHQTY = 0

            While Not adors.EOF
                mlngBAGQTY = 1
                sql = " Select TransitDaysBySea, BufferDays "
                sql = sql & "  From ScheduleParameter_mst"
                sql = sql & " Where Customer_code ='" & Trim(txtCustomerCode.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                rsTransitDays = New ADODB.Recordset
                rsTransitDays.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
                If rsTransitDays.EOF = False Then
                    Transit_Days = IIf(IsDBNull(rsTransitDays.Fields("TransitDaysBySea").Value), 0, rsTransitDays.Fields("TransitDaysBySea").Value)
                    Buffer_Days = IIf(IsDBNull(rsTransitDays.Fields("BufferDays").Value), 0, rsTransitDays.Fields("BufferDays").Value)
                End If
                If rsTransitDays.State Then rsTransitDays.Close()

                rsbagqty.GetResult("select bag_qty from item_mst where item_code = '" & adors.Fields("Item_Code").Value & "' and Status = 'A' AND UNIT_CODE = '" & gstrUNITID & "'")
                If rsbagqty.GetNoRows > 0 Then
                    mlngBAGQTY = rsbagqty.GetValue("bag_qty")
                End If

                Rs = New ADODB.Recordset
                If Darwin_FileType = "EDIFACT" Then
                    sql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST  WHERE CUST_VENDOR_CODE = '" & adors.Fields("PARTY_ID1").Value & "' AND DOCK_CODE = '" & adors.Fields("PARTY_ID3").Value & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                    Rs.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                End If

                If Darwin_FileType = "COVISINT" Then
                    sql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST WHERE CUST_VENDOR_CODE = '" & adors.Fields("PARTY_ID1").Value & "' AND DOCK_CODE = '" & adors.Fields("Factory_Code").Value & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                    Rs.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                End If

                mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                sql = " select max(dt) as dt from Calendar_Mfg_mst" & _
                    " where work_flg=0 and dt < = '" & DateAdd("d", -(Transit_Days + Buffer_Days), adors.Fields("DELIVERY_DT").Value) & "'" & _
                    " AND UNIT_CODE = '" & gstrUNITID & "'"
                rsDate = New ADODB.Recordset
                rsDate.Open(sql, mP_Connection)
                If Darwin_FileType = "EDIFACT" Then
                    If adors.Fields("Frequency").Value = "" And adors.Fields("DISPATCH_PATTERN").Value = "" Then
                        GoTo SKIP
                    End If
                End If

                If DateDiff(Microsoft.VisualBasic.DateInterval.Day, rsDate.Fields("dt").Value, ServerDate()) > 0 Or IIf(IsDBNull(adors.Fields("shipqty").Value), 0, adors.Fields("shipqty").Value) = 0 Then
                    GoTo SKIP
                End If


                SCHQTY = SCHQTY + adors.Fields("shipqty").Value

                rsbagqty.GetResult("select bag_qty from item_mst where item_code = '" & adors.Fields("Item_Code").Value & "' and Status = 'A'" & _
                                    " AND UNIT_CODE = '" & gstrUNITID & "'")
                mlngBAGQTY = 1
                If rsbagqty.GetNoRows > 0 Then
                    mlngBAGQTY = Val(rsbagqty.GetValue("bag_qty"))
                Else
                    mlngBAGQTY = 1
                End If

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

SKIP:
                mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(" INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
                    " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
                    " Shipment_Qty,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
                    " Updt_Dt,Updt_Uid,CallOffNoResetRemarks,dailypullflag,CONSIGNEE_CODE,BAG_QTY,transitDays, bufferDays,UNIT_CODE)" & _
                    " VALUES('" & Trim(TRANS_NUMBER) & "','" & adors.Fields("PARTY_ID1").Value & "', " & _
                    " '" & Trim(adors.Fields("Cust_DrgNo").Value) & "','" & Trim(adors.Fields("DELIVERY_DT").Value) & "'," & _
                    " '" & Val(adors.Fields("shipqty").Value) & "','" & getDateForDB(rsDate.Fields("dt").Value) & "'," & _
                    " '" & SCHQTY & "',0,0," & _
                    " '" & Val(adors.Fields("shipqty").Value) & "',getDate(),'" & mP_User & "',getDate()," & _
                    " '" & mP_User & "','" & Replace(Remarks, "'", "''") & "','0'," & _
                    " '" & txtCustomerCode.Text & "','" & mlngBAGQTY & "'," & Transit_Days & "," & Buffer_Days & ",'" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)  ''eMpro -20080528 - 19328 : TXTCUSTOMER REPLACED WITH RS!CUSTOMER_CODE
                rsDate.Close()

                SCHQTY = 0

                adors.MoveNext()

                If Rs.State Then Rs.Close()
                Rs = Nothing
            End While

            If adors.State Then adors.Close()
            adors = Nothing
        Else
            MsgBox("Schedule Upload Required Flag is OFF for Customer " + txtCustomerCode.Text + vbCrLf + "Can't Upload File...")
            If adors.State Then adors.Close()
            adors = Nothing
            Return Nothing
            Exit Function
        End If
        Return Nothing
        Exit Function
Err:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        If adors.State = ADODB.ObjectStateEnum.adStateOpen Then adors.Close() : adors = Nothing
        If Rs.State = ADODB.ObjectStateEnum.adStateOpen Then Rs.Close() : Rs = Nothing
        RSWHSTOCK.ResultSetClose() : RSWHSTOCK = Nothing

        Return Nothing
        Exit Function

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
        sql = "select TOP 1 BackUpLocation from scheduleparameter_mst" & " WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' ORDER BY UPDDT DESC "
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

                    FSO.MoveFile(folderName & "\" & subFileName, bkpLocation & "\")
                    mP_Connection.Execute("INSERT INTO BACKUPFILEHISTORY(" & " FILENAME, FILEDATE, STATUS,UNIT_CODE)" & " VALUES('" & filearray(UBound(filearray)) & "'," & " getDate(),'" & status & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Else
                    MsgBox("Source path does not exist", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            Next upldFiles

            lblMessage3.Text = "Transaction completed Successfully."
            mblnfilemove = True
            mblnDailymktUpdated = True
        End If
        Return Nothing
        Exit Function
ERR_Renamed:
        If Err.Number = 70 Then
            lblMessage3.Text = "Backup Location is ReadOnly."
            Exit Function
        End If
        If Err.Number = 76 Then
            lblMessage3.Text = "BackUp Location Not Found."
            Exit Function
        End If
        If Err.Number = 5 Then
            lblMessage3.Text = "File Already Open, Cann't Move."
            Exit Function
        End If

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        Exit Function

    End Function

    Private Sub FN_FILESELECTION()
        On Error GoTo ErrHandler
        Dim folderName As String = Nothing
        Dim upldFiles As Scripting.File
        Dim filearray(0) As Object
        Dim upldFileName(0) As Object
        Dim filedate(0) As Object
        Dim latestFile As String = ""
        Dim YesNo As String = Nothing
        Dim Temp As String = Nothing
        Dim Rs As New ClsResultSetDB
        Dim RSMISSINGCALLOFFREQD As New ClsResultSetDB

        'STOP UPLOADING IF MAX "AS ON DATE" OF  WAREHOUSE STOCK STATUS IS NOT SAME FOR ALL THE CONSIGNEES OF A CUSTOMER

        Dim rsWHConsignee As ADODB.Recordset
        Dim strsql As String = String.Empty
        Dim strConsList As String = String.Empty
        Dim strMaxWhDate As String = String.Empty
        Dim rsSHipmentthruWH As ADODB.Recordset
        Dim blnShipmentThruWh As Boolean
        Dim blnWHStockReqdForAllConsignee As Boolean

        rsSHipmentthruWH = New ADODB.Recordset
        rsSHipmentthruWH.Open("Select ShipmentThruWh from Customer_mst where Customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        blnShipmentThruWh = rsSHipmentthruWH.Fields("ShipmentThruWH").Value
        rsSHipmentthruWH.Close()
        rsSHipmentthruWH = Nothing

        rsSHipmentthruWH = New ADODB.Recordset
        rsSHipmentthruWH.Open("Select WHStockReqdForAllConsignee from sales_parameter where UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        blnWHStockReqdForAllConsignee = rsSHipmentthruWH.Fields("WHStockReqdForAllConsignee").Value
        rsSHipmentthruWH.Close()
        rsSHipmentthruWH = Nothing

        If blnShipmentThruWh = True And blnWHStockReqdForAllConsignee = True Then
            strsql = " select max(trans_dt) AS TRANS_DT from warehouse_stock_dtl where customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'"

            rsWHConsignee = New ADODB.Recordset
            rsWHConsignee.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            If rsWHConsignee.RecordCount > 0 Then

                strMaxWhDate = VB6.Format(rsWHConsignee.Fields("trans_dt").Value, "dd MMM yyyy")

                strsql = "select distinct w.consignee_code from warehouse_stock_dtl w" & _
                    " inner join CustWarehouse_Mst C on" & _
                    " c.Customer_Code = w.Consignee_Code and " & _
                    " c.UNIT_CODE = w.UNIT_CODE and" & _
                    " c.WH_code = w.WareHouse_Code" & _
                    " where c.Active = 1 and " & _
                    " w.consignee_code not in (" & _
                    " select distinct(consignee_code) from warehouse_stock_dtl" & _
                    " where customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "' and trans_dt = (" & _
                    " select max(trans_dt) from warehouse_stock_dtl where customer_code = '" & txtCustomerCode.Text & "'" & _
                    " AND UNIT_CODE = '" & gstrUNITID & "')) AND w.CUSTOMER_CODE = '" & txtCustomerCode.Text & "'" & _
                    " AND w.UNIT_CODE = '" & gstrUNITID & "'"

                If rsWHConsignee.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsWHConsignee.Close()
                    rsWHConsignee = Nothing
                End If
                rsWHConsignee = New ADODB.Recordset
                rsWHConsignee.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

                If rsWHConsignee.RecordCount > 0 Then
                    rsWHConsignee.MoveFirst()
                    While Not rsWHConsignee.EOF
                        strConsList = strConsList + rsWHConsignee.Fields("consignee_code").Value + vbCrLf
                        rsWHConsignee.MoveNext()
                    End While
                End If

                If Trim(strConsList).Length > 0 Then
                    MsgBox("Warehouse Stock Not Uploaded For Following Consignees Of " & vbCrLf & UCase(txtCustomerCode.Text) & " On " & strMaxWhDate & vbCrLf & vbCrLf & strConsList & vbCrLf & "Upload Warehouse Stock For These Consignees.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If

                If rsWHConsignee.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsWHConsignee.Close()
                    rsWHConsignee = Nothing
                End If

            End If
        End If

        RSMISSINGCALLOFFREQD.GetResult("SELECT MissingCallOff_Reqd FROM SALES_PARAMETER WHERE  UNIT_CODE = '" & gstrUNITID & "'")

        If OptReleaseFile.Checked Then
            Rs.GetResult("select top 1 ReleaseFile_Location" & " from scheduleparameter_mst" & " where customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'" & " order by entdt")
            txtFileName.Text = Rs.GetValue("ReleaseFile_Location")
            Rs.ResultSetClose()
        End If

        If Len(LTrim(RTrim(txtFileName.Text))) > 0 Then
            folderName = txtFileName.Text

            Temp = Mid(StrReverse(txtFileName.Text), 1, InStr(1, StrReverse(txtFileName.Text), "\") - 1)
            Obj_FSO = New Scripting.FileSystemObject
            If InStr(1, Temp, ".") = 0 Then
                If Obj_FSO.FolderExists(folderName) = False Then
                    MsgBox("Folder Does Not Exist")
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If

                    Exit Sub
                End If
            Else
                If Obj_FSO.FileExists(txtFileName.Text) = False Then
                    MsgBox("No Call-Offs present in the Release Folder.")
                    Obj_FSO = Nothing
                    If Not Obj_EX Is Nothing Then
                        KillExcelProcess(Obj_EX)
                        Obj_EX = Nothing
                    End If

                    Exit Sub
                End If
                folderName = VB.Left(folderName, Len(folderName) - Len(Temp) - 1)
            End If

            If Obj_FSO.GetFolder(folderName).Files.Count > 0 Then
                If Obj_FSO.GetFolder(folderName).Files.Count > 1 Then

                    For Each upldFiles In Obj_FSO.GetFolder(folderName).Files
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
                    Next upldFiles

                    latestFile = filearray(UBound(filearray))
                Else
                    For Each upldFiles In Obj_FSO.GetFolder(folderName).Files
                        latestFile = upldFiles.Path
                        latestFile = StrReverse(Mid(StrReverse(latestFile), 1, InStr(1, StrReverse(latestFile), "\") - 1))
                    Next upldFiles
                End If

            End If

            If Obj_FSO.GetFolder(folderName).Files.Count > 1 Then
                MsgBox("You Have More Than One File To Upload...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
            If Trim(latestFile) <> "" Then
                txtFileName.Text = Obj_FSO.GetFolder(folderName).Path & "\" & latestFile ''& ".csv"
            End If

            If Obj_FSO.FileExists(txtFileName.Text) = False Then
                MsgBox(" No File Exists in the Release Folder.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            Call FN_Release_File_Upload()

        End If

        If Darwin_FileType <> "VDA" And Darwin_FileType <> "EDIFACT" And Darwin_FileType <> "COVISINT" And Darwin_FileType <> "BOSCH" And Darwin_FileType <> "PKC" Then
            MsgBox("Wrong File Type-It's not VDA/EDIFACT/COVISINT format.", MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If

        Exit Sub
ErrHandler:
        If Err.Number = 5 Then
            MsgBox("Invalid File Name.", MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Function FN_DAILYPULLQTY(ByVal LNGWHSTOCK As Object, ByVal lngISSUEDQTY As Object, ByVal lngRCVDQTY As Object, ByVal lngSCHQTY As Long, ByVal varWHCODE As Object, ByVal lngSAFETYDAYS As Long, ByVal StrItemCode As String, ByVal strCustDrgNo As String, ByVal Row As Integer, ByVal strFACTORY_CODE As String, ByVal mlngBAGQTY As Long) As String
        On Error GoTo ErrHandler
        'FN_DAILYPULLQTY(varWhStock, varIssuedQty, varRcvdQty, SCHQTY, Adors!SafetyDays, varWHCODE, CStr(Adors!Item_Code), CInt(Row), CStr(Adors!FactoryCode))
        'Created By         : Shubhra Verma
        'Created On         : 04 Mar 2008 to 07 Mar 2008
        'Issue id           : eMpro-20080306-13517
        'Revision History   :1 - There should be a provision of using daily pull
        '                    qty from Warehouse_Stock_dtl as minimum
        '                    safety stock if daily pull qty check box is checked
        '                    in CDP form.

        Dim rsbagqty As New ClsResultSetDB
        Dim SCHQTY As Long
        Dim lngdlypullqty As Long
        Dim SQLDailyPullQty As SqlCommand
        Dim rdrDailyPullQty As SqlDataReader
        Dim strRetString As String

        SQLDailyPullQty = New SqlCommand
        SQLDailyPullQty.Connection = SqlConnectionclass.GetConnection

        SQLDailyPullQty.CommandText = "SELECT  RATE FROM WAREHOUSE_STOCK_DTL" & _
        " WHERE CUSTOMER_CODE = '" & txtCustomerCode.Text & "'" & _
        " AND WAREHOUSE_CODE = '" & varWHCODE & "'" & _
        " and item_code = '" & strCustDrgNo & "' AND UNIT_CODE = '" & gstrUNITID & "'" & _
        " and  TRANS_DT  = (SELECT MAX(TRANS_DT)FROM WAREHOUSE_STOCK_DTL" & _
        " WHERE  WAREHOUSE_CODE = '" & varWHCODE & "'" & _
        " and customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "')" & _
        " and consignee_code in (select customer_code from customer_mst" & _
        " where dock_code = '" & strFACTORY_CODE & "'" & _
        " and cust_vendor_code = '" & varWHCODE & "' AND UNIT_CODE = '" & gstrUNITID & "')" & _
        " and revno =  (select max(revno)" & _
        " From warehouse_stock_dtl" & _
        " WHERE CUSTOMER_CODE='" & txtCustomerCode.Text & "'" & _
        " AND WAREHOUSE_CODE='" & varWHCODE & "' AND UNIT_CODE = '" & gstrUNITID & "'" & _
        " and consignee_code in (select customer_code from customer_mst" & _
        " where dock_code = '" & strFACTORY_CODE & "'" & _
        " and cust_vendor_code = '" & varWHCODE & "' AND UNIT_CODE = '" & gstrUNITID & "')" & _
        " and TRANS_DT  = (SELECT MAX(TRANS_DT)FROM WAREHOUSE_STOCK_DTL" & _
        " WHERE  WAREHOUSE_CODE = '" & varWHCODE & "'" & _
        " and customer_code = '" & txtCustomerCode.Text & "' AND UNIT_CODE = '" & gstrUNITID & "'))"

        rdrDailyPullQty = SQLDailyPullQty.ExecuteReader

        strRetString = "0"

        If rdrDailyPullQty.HasRows Then
            rdrDailyPullQty.Read()
            lngdlypullqty = IIf(rdrDailyPullQty("RATE").ToString = "", 0, rdrDailyPullQty("RATE").ToString)
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

            If rdrDailyPullQty.HasRows Then
                strRetString = CStr(SCHQTY) + "*" + CStr(IIf(rdrDailyPullQty("RATE").ToString = "", 0, rdrDailyPullQty("RATE").ToString))
            Else
                strRetString = CStr(SCHQTY) + "*" + "0"
            End If
            rdrDailyPullQty.Close()
            SQLDailyPullQty.Dispose()
            SQLDailyPullQty = Nothing

        End If

        FN_DAILYPULLQTY = strRetString
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        If Not SQLDailyPullQty Is Nothing Then
            SQLDailyPullQty.Dispose()
            SQLDailyPullQty = Nothing
        End If

        Obj_EX.Workbooks.Close()
        If Not Obj_EX Is Nothing Then
            KillExcelProcess(Obj_EX)
            Obj_EX = Nothing
        End If

    End Function

    Private Sub txtNoOfMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtNoOfMonths.Validating
        On Error GoTo ErrHandler
        If optAvgofNextMonths.Enabled = True And optAvgofNextMonths.Visible = True Then
            If Val(txtNoOfMonths.Text) <= 1 Then
                MsgBox("No Of Months Must Be Greater Than 1", MsgBoxStyle.Information, ResolveResString(100))
                txtNoOfMonths.Focus()
            End If
        End If
        Exit Sub

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FN_TRANSFERDATAINCOVISINT(ByVal DOC_NO As String, ByVal FILETYPE As String)
        On Error GoTo ErrHandler
        Dim rsTMP_DAYWISESCHEDULE As ClsResultSetDB
        Dim strSql As String = String.Empty
        Dim rsInsertHdr As ClsResultSetDB
        Dim rsInsertDtl As ClsResultSetDB

        If FILETYPE = "VDA" Then
            strSql = "INSERT INTO TMP_DAYWISESCHEDULE" & _
                " SELECT	A.DOC_NO,COUNT(DT) WORKINGDAYS,MONTH(DT) MONTH,YEAR(DT) YEAR,A.QTY," & _
                " ROUND(A.QTY / COUNT(DT),0,1) AS QTYPERDAY," & _
                " A.QTY % COUNT(DT) REMAININGQTY," & _
                " A.CUST_DRGNO,A.ITEM_CODE,A.CONSIGNEE_CODE," & _
                " A.GI_VEND_CODE, A.FACTORYCODE, '" & gstrUNITID & "'" & _
                " FROM CALENDAR_MKT_CUST C," & _
                " ( " & _
                " SELECT MONTH(V.DDRD_Req_Dt1) MONTH,YEAR(V.DDRD_Req_Dt1) YEAR," & _
                " SUM(V.DDRD_REQ_QTY1) QTY,  " & _
                " V.CUST_DRGNO,V.ITEM_CODE,V.CONSIGNEE_CODE," & _
                " V.GI_VEND_CODE, V.FACTORYCODE, V.DOC_NO" & _
                " FROM VW_SCHEDULE_PROPOSAL V" & _
                " WHERE(V.DOC_NO = " & DOC_NO & " and V.UNIT_CODE = '" & gstrUNITID & "')" & _
                " GROUP BY MONTH(DDRD_Req_Dt1),YEAR(DDRD_Req_Dt1),CUST_DRGNO," & _
                " ITEM_CODE, CONSIGNEE_CODE, GI_VEND_CODE, FACTORYCODE, V.DOC_NO" & _
                " )A" & _
                " WHERE	MONTH(DT)	=	A.MONTH AND" & _
                " YEAR(DT)	=	A.YEAR  AND" & _
                " C.WORK_FLG = 0 AND C.CUST_CODE = A.CONSIGNEE_CODE AND C.DT >= CONVERT(VARCHAR(12),GETDATE(),106) AND C.UNIT_CODE = '" & gstrUNITID & "'" & _
                " GROUP BY A.DOC_NO,MONTH(DT) ,YEAR(DT),A.CUST_DRGNO,A.ITEM_CODE,A.CONSIGNEE_CODE," & _
                " A.GI_VEND_CODE,A.FACTORYCODE ,A.QTY"

            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            strSql = String.Empty
            strSql = "INSERT INTO SCHEDULE_UPLOAD_COVISINT_HDR" & _
                " SELECT DISTINCT V.DOC_NO,'302',V.CUST_CODE,H.plant_c,H.UPLOAD_FILE_NAME,H.upload_file_type,GETDATE(),'" & mP_User & "',getdate()," & _
                " '" & mP_User & "',V.CONSIGNEE_CODE,getdate(), '" & gstrUNITID & "'" & _
                " FROM VW_SCHEDULE_PROPOSAL V, SCHEDULE_UPLOAD_DARWIN_HDR H" & _
                " WHERE V.DOC_NO = " & DOC_NO & " AND V.DOC_NO = H.DOC_NO AND V.UNIT_CODE = H.UNIT_CODE AND V.UNIT_CODE = '" & gstrUNITID & "'"

            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            strSql = String.Empty
            strSql = "INSERT INTO SCHEDULE_UPLOAD_COVISINT_DTL" & _
                " SELECT D.DOC_NO,'302',D.CUST_DRGNO,D.GI_VEND_CODE,D.FACTORYCODE," & _
                " D.CONSIGNEE_CODE,C.DT,D.QTYPERDAY,GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "', '" & gstrUNITID & "'" & _
                " FROM CALENDAR_MKT_CUST C, TMP_DAYWISESCHEDULE D" & _
                " WHERE MONTH(C.DT) = D.MONTH AND" & _
                " Year(C.DT) = D.YEAR" & _
                " AND C.WORK_FLG = 0 and C.UNIT_CODE = D.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'" & _
                " AND D.DOC_NO = " & DOC_NO & " AND C.DT >= CONVERT(VARCHAR(12),GETDATE(),106) AND D.CONSIGNEE_CODE = C.CUST_CODE"

            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            mP_Connection.Execute("delete from schedule_upload_darwin_hdr where doc_no = " & DOC_NO & " AND UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute("delete from schedule_upload_darwin_dtl where doc_no = " & DOC_NO & " AND UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        End If

        If FILETYPE = "EDIFACT" Then
            strSql = String.Empty
            strSql = "INSERT INTO TMP_DAYWISESCHEDULE" & _
              " SELECT	A.DOC_NO,COUNT(DT) WORKINGDAYS,MONTH(DT) MONTH,YEAR(DT) YEAR,A.QTY," & _
              " ROUND(A.QTY / COUNT(DT),0,1) AS QTYPERDAY," & _
              " A.QTY % COUNT(DT) REMAININGQTY," & _
              " A.CUST_DRGNO,A.ITEM_CODE,A.CONSIGNEE_CODE,A.PARTY_ID1,'' AS FACTORYCODE, '" & gstrUNITID & "'" & _
              " FROM CALENDAR_MKT_CUST C," & _
              " (  " & _
              " SELECT MONTH(D.DELIVERY_DT) MONTH,YEAR(D.DELIVERY_DT) YEAR," & _
              " SUM(D.QUANTITY) QTY," & _
              " D.ITEM_CODE AS CUST_DRGNO,'' AS ITEM_CODE,H.CUST_CODE AS CONSIGNEE_CODE,H.PARTY_ID1,'' AS FACTORYCODE " & _
              " , D.DOC_NO" & _
              " FROM SCHEDULE_UPLOAD_DARWINEDIFACT_DTL D, SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H" & _
              " WHERE H.DOC_NO = D.DOC_NO AND H.UNIT_CODE = D.UNIT_CODE" & _
              " AND D.DOC_NO = " & DOC_NO & " AND D.UNIT_CODE = '" & gstrUNITID & "'" & _
              " GROUP BY MONTH(D.DELIVERY_DT),YEAR(D.DELIVERY_DT),D.ITEM_CODE," & _
              " H.PARTY_ID1, D.DOC_NO, H.CUST_CODE" & _
              " )A" & _
              " WHERE	MONTH(DT)	=	A.MONTH AND" & _
              " YEAR(DT)	=	A.YEAR  AND C.UNIT_CODE = '" & gstrUNITID & "' AND" & _
              " C.WORK_FLG = 0 AND C.CUST_CODE = A.CONSIGNEE_CODE AND C.DT >= CONVERT(VARCHAR(12),GETDATE(),106)" & _
              " GROUP BY A.DOC_NO,MONTH(DT) ,YEAR(DT),A.ITEM_CODE,A.PARTY_ID1,A.QTY,A.CUST_DRGNO,A.CONSIGNEE_CODE"

            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            strSql = String.Empty
            strSql = "INSERT INTO SCHEDULE_UPLOAD_COVISINT_HDR" & _
                " SELECT DISTINCT H.DOC_NO,H.DOC_TYPE,H.CUST_CODE,H.PLANT_C,H.UPLOAD_FILE_NAME, " & _
                " H.UPLOAD_FILE_TYPE,GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "',H.CUST_CODE,getdate(), '" & gstrUNITID & "'" & _
                " FROM SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H" & _
                " WHERE H.DOC_NO = " & DOC_NO & " AND H.UNIT_CODE = '" & gstrUNITID & "'"

            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            strSql = String.Empty
            strSql = "INSERT INTO SCHEDULE_UPLOAD_COVISINT_DTL" & _
                " SELECT DISTINCT D.DOC_NO,D.DOC_TYPE,D.ITEM_CODE,T.GI_VEND_CODE,'' FACTORY_CODE,T.CONSIGNEE_CODE,C.DT," & _
                " T.QTYPERDAY,GETDATE() ENT_DT,'" & mP_User & "' ENT_UID,GETDATE() UPDT_DT,'" & mP_User & "' AS UPDT_UID , '" & gstrUNITID & "'" & _
                " FROM CALENDAR_MKT_CUST C, SCHEDULE_UPLOAD_DARWINEDIFACT_DTL D, TMP_DAYWISESCHEDULE T" & _
                " WHERE MONTH(C.DT) = T.MONTH AND" & _
                " YEAR(C.DT) = T.YEAR AND C.UNIT_CODE = T.UNIT_CODE AND" & _
                " MONTH(D.DELIVERY_DT) = T.MONTH AND" & _
                " Year(D.DELIVERY_DT) = T.YEAR AND D.UNIT_CODE = T.UNIT_CODE" & _
                " AND C.WORK_FLG = 0" & _
                " AND D.ITEM_CODE = T.CUST_DRGNO" & _
                " AND D.DOC_NO = T.DOC_NO" & _
                " AND T.DOC_NO = " & DOC_NO & " AND C.DT >= CONVERT(VARCHAR(12),GETDATE(),106) AND T.CONSIGNEE_CODE = C.CUST_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'"

            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        End If

        strSql = String.Empty
        strSql = "update schedule_upload_covisint_dtl" & _
            " set schedule_upload_covisint_dtl.qty = tmp_daywiseschedule.qtyperday + tmp_daywiseschedule.remainingqty " & _
            " from schedule_upload_covisint_dtl, tmp_daywiseschedule" & _
            " where schedule_upload_covisint_dtl.delivery_date in (" & _
            " Select min(delivery_date) " & _
            " from schedule_upload_covisint_dtl" & _
            " where DOC_NO = " & DOC_NO & " AND UNIT_CODE = '" & gstrUNITID & "'" & _
            " group by month(delivery_date),year(delivery_date))" & _
            " and month(delivery_date) = tmp_daywiseschedule.month and year(delivery_date) = tmp_daywiseschedule.year" & _
            " and schedule_upload_covisint_dtl.doc_no = tmp_daywiseschedule.doc_no AND schedule_upload_covisint_dtl.UNIT_CODE = tmp_daywiseschedule.UNIT_CODE" & _
            " and tmp_daywiseschedule.doc_no = " & DOC_NO & " AND schedule_upload_covisint_dtl.UNIT_CODE = '" & gstrUNITID & "' and schedule_upload_covisint_dtl.item_code = tmp_daywiseschedule.cust_drgno"

        mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        mP_Connection.Execute("delete from schedule_upload_darwinEdifact_hdr where doc_no = " & DOC_NO & " AND UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute("delete from schedule_upload_darwinEdifact_dtl where doc_no = " & DOC_NO & " AND UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub ChkTextFile_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkTextFile.CheckedChanged
        On Error GoTo ErrHandler

        If ChkTextFile.Checked = False Then
            ChkDaimler.Checked = False
            ChkFord.Checked = False
            ChkDaimler.Enabled = False
            ChkFord.Enabled = False
        End If

        If ChkTextFile.Checked = True Then
            ChkDaimler.Checked = True
            ChkFord.Checked = False
            ChkDaimler.Enabled = True
            ChkFord.Enabled = True
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub ChkDaimler_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkDaimler.CheckedChanged
        On Error GoTo ErrHandler

        If ChkDaimler.Checked = True Then
            ChkFord.Checked = False
        End If
        If ChkDaimler.Checked = False Then
            ChkFord.Checked = True
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub ChkFord_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkFord.CheckedChanged
        On Error GoTo ErrHandler

        If ChkFord.Checked = True Then
            ChkDaimler.Checked = False
        End If
        If ChkFord.Checked = False Then
            ChkDaimler.Checked = True
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub BackDateSchedule(ByVal TRANS_NUMBER As Integer, ByVal Cust_DrgNo As String, ByVal varWhStock As String, ByVal CONSIGNEE_CODE As String, ByVal WHDATE As Date, ByVal varRevNo As Integer, ByVal blnDAILYPULLFLAG As Boolean, ByVal safetydaysmin As Integer, ByVal safetydaysMAX As Integer, ByVal SafetyDays As Integer, ByVal STOCKCALCWADAYS As Char, ByVal dailypullrate As Integer, ByVal ScheduleCalcMonths As Integer, ByVal DAYSFORSAFETYSTOCK As Integer, ByVal Transit_Days As Integer, ByVal Buffer_Days As Integer)
        Dim strSQL As String = String.Empty

        'Try
        '    If Darwin_FileType = "VDA" Then

        '        strSQL = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
        '            " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
        '            " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
        '            " Updt_Dt,Updt_Uid,WH_DATE,REVNO,CALLOFFNORESETREMARKS,DAILYPULLFLAG" & _
        '            " ,SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
        '            " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays, bufferDays,UNIT_CODE)" & _
        '            " VALUES('" & Trim(TRANS_NUMBER) & "',getdate(), " & _
        '            " '" & Cust_DrgNo & "',getdate()," & _
        '            " '" & Val(varWhStock) & "',getdate()," & _
        '            " '" & Val(varWhStock) & "','" & CONSIGNEE_CODE & "','" & varWhStock & "'," & _
        '            " 0,'" & Val(varWhStock) & "',getDate(),'" & mP_User & "',getDate(),'" & mP_User & "'," & _
        '            " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "' , '" & blnDAILYPULLFLAG & "'," & _
        '            " " & safetydaysmin & "," & safetydaysMAX & "," & _
        '            " " & SafetyDays & "," & mlngBAGQTY & "," & _
        '            " '" & STOCKCALCWADAYS & "','" & ScheduleCalcMonths & "'" & _
        '            " ,'" & DAYSFORSAFETYSTOCK & "'," & Val(dailypullrate) & ",'" & varWhStock & "'," & Transit_Days & "," & _
        '            " " & Buffer_Days & ",'" & gstrUNITID & "')"
        '        sqlInsertUpdate.ExecuteNonQuery()
        '    End If

        '    If FileType = "EDIFACT" Then
        '        sqlInsertUpdate.CommandText = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
        '                " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
        '                " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
        '                " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG," & _
        '                " SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
        '                " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY,transitDays,bufferDays,UNIT_CODE)" & _
        '                " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("PARTY_ID1").ToString & "', " & _
        '                " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "','" & Format(CDate(sqlRdrDisplay("DELIVERY_DT").ToString), "dd MMM yyyy") & "'," & _
        '                " '" & Val(sqlRdrDisplay("shipqty").ToString) & "','" & Format(dtDate, "dd MMM yyyy") & "'," & _
        '                " '" & SCHQTY & "','" & CONSIGNEE_CODE & "','" & IIf(IsDBNull(varWhStock), 0, varWhStock) & "'," & _
        '                " '" & varRcvdQty & "','" & varIssuedQty & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
        '                " '" & Convert.ToDateTime(WHDATE).ToString("dd MMM yyyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
        '                " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
        '                " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
        '                " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
        '                " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
        '                " " & Buffer_Days & ",'" & gstrUNITID & "' )"
        '        sqlInsertUpdate.ExecuteNonQuery()
        '    End If

        '    If FileType = "COVISINT" Then
        '        sqlInsertUpdate.CommandText = "INSERT INTO SCHEDULEPROPOSALCALCULATIONS ( " & _
        '            " Doc_No,Wh_Code,Item_Code,release_Dt,release_Qty,Shipment_Dt, " & _
        '            " Shipment_Qty,consignee_code,Wh_Stock,received_Qty,Issued_Qty,Ent_Dt,Ent_Uid," & _
        '            " Updt_Dt,Updt_Uid,wh_date,revno,CALLOFFNORESETREMARKS,DAILYPULLFLAG" & _
        '            " ,SAFETYDAYSMIN,SAFETYDAYSMAX,SAFETYDAYS,BAG_QTY,STOCKCALCWADAYS," & _
        '            " SCHEDULECALCMONTHS,DAYSFORSAFETYSTOCK,DAILYPULLRATE,SUMOFRELEASEQTY, transitDays,bufferDays,UNIT_CODE)" & _
        '            " VALUES('" & Trim(TRANS_NUMBER) & "','" & sqlRdrDisplay("wh_code").ToString & "', " & _
        '            " '" & Trim(sqlRdrDisplay("Cust_DrgNo").ToString) & "','" & Format(CDate(sqlRdrDisplay("DELIVERY_DATE").ToString), "dd MMM yyyy") & "'," & _
        '            " '" & Val(sqlRdrDisplay("shipqty").ToString) & "','" & Format(dtDate, "dd MMM yyy") & "'," & _
        '            " '" & SCHQTY & "','" & sqlRdrDisplay("CONSIGNEE_CODE").ToString & "','" & varWhStock & "'," & _
        '            " '" & varRcvdQty & "','" & varIssuedQty & "',getDate(),'" & mP_User & "' ,getDate(),'" & mP_User & "'," & _
        '            " '" & Format(WHDATE, "dd MMM yyy") & "','" & varRevNo & "','" & Replace(Remarks, "'", "''") & "', '" & blnDAILYPULLFLAG & "'," & _
        '            " " & sqlRdrDisplay("safetydaysmin").ToString & "," & sqlRdrDisplay("safetydaysMAX").ToString & "," & _
        '            " " & sqlRdrDisplay("SAFETYDAYS").ToString & "," & mlngBAGQTY & "," & _
        '            " '" & Val(sqlRdrDisplay("STOCKCALCWADAYS").ToString) & "','" & Val(sqlRdrDisplay("ScheduleCalcMonths").ToString) & "'" & _
        '            " ,'" & Val(sqlRdrDisplay("DAYSFORSAFETYSTOCK").ToString) & "'," & Val(dailypullrate) & ",'" & Val(sqlRdrDisplay("SUMOFRELEASEQTY").ToString) & "'," & Transit_Days & "," & _
        '            " " & Buffer_Days & ",'" & gstrUNITID & "' )"
        '        sqlInsertUpdate.ExecuteNonQuery()
        '    End If
        'Catch ex As Exception
        '    Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        'End Try
    End Sub

    Private Sub FN_Release_File_Upload_COVISINT()
        Try

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

End Class