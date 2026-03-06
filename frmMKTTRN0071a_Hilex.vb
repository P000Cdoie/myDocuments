Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Friend Class frmMKTTRN0071a_Hilex
    Inherits System.Windows.Forms.Form
    '-------------------------------------------------------------------------------------------
    ' Modified By       :   Prashant Rajpal
    ' Modified On       :   27th-28th  July 2011
    ' Issue id          :   10119746  
    ' Purpose           :   Sales Order Filteration incorporated
    '-------------------------------------------------------------------------------------------
    'Modified By	:   Ekta Uniyal
    'Modified On	:   4 mar 2014
    'Description	:   To support multi-unit functionality for Hilex.

    'REVISED BY     :  ASHISH SHARMA
    'REVISED DATE   :  08-06-2017
    'ISSUE ID       :  101188073 
    'PURPOSE        :  GST CHANGES
    '-------------------------------------------------------------------------------------------	
    Dim mstrInvType As String
    Dim mstrInvSubType As String
    Dim mblnBackDatePDS As Boolean
    Dim mstartingdate As Date
    '101188073 Start
    Dim _itemCode As String
    '101188073 End
    Dim mblnMultipleSOPDS As Boolean
    Dim mstraccountcode As String
    Enum GridHeader
        Mark = 1
        ItemCode = 2
        DrawingNo = 3
        Description = 4
        Quantity = 5
        DespatchQty = 6
        CustRef = 7
        AmendmentNo = 8
        SchDate = 9
        SChTime = 10
        AccountCode = 11
        'PDSNo = 12
        FTSITEM = 12
        FTSBARCODE = 13
        '101188073 Start
        IS_HSN_SAC = 14
        HSN_SAC_CODE = 15
        CGST_TYPE = 16
        SGST_TYPE = 17
        IGST_TYPE = 18
        UTGST_TYPE = 19
        COMPENSATION_CESS_TYPE = 20
        '101188073 End
    End Enum

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error GoTo ErrHandler
        Me.Dispose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        On Error GoTo ErrHandler
        mstrItemText = ""
        Dim intSubItem As Short
        Dim StrItemCode As String
        Dim strDrawingNo As String
        Dim strDescription As String
        Dim dblQuantity As Double
        Dim dbldespatchQty As Double
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim strSchDate As String
        Dim strSchTime As String
        Dim strAccountCode As String
        Dim strSrvDINo As String
        Dim blnftsitem As Boolean
        Dim blnftsBarcode As Boolean
        '101188073 Start
        Dim strItem As String()
        Dim flag As Boolean = False
        Dim hsnsacType As String = String.Empty
        Dim hsnsacCode As String = String.Empty
        Dim cgstType As String = String.Empty
        Dim sgstType As String = String.Empty
        Dim igstType As String = String.Empty
        Dim utgstType As String = String.Empty
        Dim CompensationCessType As String = String.Empty
        Dim cgstPercent As String = String.Empty
        Dim sgstPercent As String = String.Empty
        Dim igstPercent As String = String.Empty
        Dim utgstPercent As String = String.Empty
        Dim CompensationCessPercent As String = String.Empty
        Dim sqlCmd As SqlCommand
        Dim dsTax As DataSet
        '101188073 End

        If Not ValidateData Then Exit Sub
        '101188073 Start
        If gblnGSTUnit Then
            If Not ValidateGSTTaxes() Then
                Exit Sub
            Else
                If Len(_itemCode) > 0 Then
                    strItem = _itemCode.Split(",")
                End If
            End If
        End If
        '101188073 End
        With SpItems
            For intSubItem = 1 To .maxRows
                .Row = intSubItem
                .Col = GridHeader.Mark
                If CBool(.Value) = True Then
                    '101188073 Start
                    If gblnGSTUnit Then
                        If strItem IsNot Nothing AndAlso strItem.Count > 0 Then
                            For arrayIndex As Integer = 0 To strItem.Count - 1
                                .Row = intSubItem
                                .Col = GridHeader.ItemCode
                                If Trim(.Text) = Trim(strItem(arrayIndex).ToString()) Then
                                    flag = True
                                    Exit For
                                End If
                            Next
                            If flag Then
                                Continue For
                            End If
                        End If
                    End If
                    '101188073 End
                    .Col = GridHeader.ItemCode
                    StrItemCode = Trim(.Text)
                    .Col = GridHeader.Description
                    strDescription = Trim(.Text)
                    .Col = GridHeader.Quantity
                    dblQuantity = CDbl(Trim(.Text))
                    .Col = GridHeader.DespatchQty
                    dbldespatchQty = CDbl(Trim(.Text))
                    .Col = GridHeader.CustRef
                    strCustRef = Trim(.Text)
                    .Col = GridHeader.AmendmentNo
                    StrAmendmentNo = Trim(.Text)
                    .Col = GridHeader.SchDate
                    strSchDate = Trim(.Text)
                    .Col = GridHeader.SChTime
                    strSchTime = Trim(.Text)
                    .Col = GridHeader.AccountCode
                    strAccountCode = Trim(.Text)
                    .Col = GridHeader.DrawingNo
                    strDrawingNo = Trim(.Text)
                    strSrvDINo = Trim(Me.TxtPDSNo.Text)

                    blnftsitem = CBool(Trim(Find_Value("select FTS_ITEM from item_mst where UNIT_CODE = '" & gstrUnitId & "' AND item_code='" & StrItemCode & "'")))
                    blnftsBarcode = CBool(Trim(Find_Value("select FTS_BARCODE_TRACKING from item_mst where UNIT_CODE = '" & gstrUnitId & "' AND item_code='" & StrItemCode & "'")))
                    frmMKTTRN0071_HILEX.FTSItem = blnftsitem
                    frmMKTTRN0071_HILEX.FTSBarcode = blnftsBarcode
                    '101188073 Start
                    If gblnGSTUnit Then
                        .Row = intSubItem
                        .Col = GridHeader.IS_HSN_SAC
                        hsnsacType = .Text
                        .Col = GridHeader.HSN_SAC_CODE
                        hsnsacCode = .Text
                        .Col = GridHeader.CGST_TYPE
                        cgstType = .Text
                        .Col = GridHeader.SGST_TYPE
                        sgstType = .Text
                        .Col = GridHeader.IGST_TYPE
                        igstType = .Text
                        .Col = GridHeader.UTGST_TYPE
                        utgstType = .Text
                        .Col = GridHeader.COMPENSATION_CESS_TYPE
                        CompensationCessType = .Text
                        sqlCmd = New SqlCommand("SELECT CGST_PERCENT,SGST_PERCENT,IGST_PERCENT,UTGST_PERCENT,COMPENSATION_CESS_PERCENT FROM dbo.UDF_GST_TAX_RATE_PERCENT('" & gstrUnitId & "','" & cgstType & "','" & sgstType & "','" & igstType & "','" & utgstType & "','" & CompensationCessType & "')")
                        sqlCmd.CommandType = CommandType.Text
                        dsTax = SqlConnectionclass.GetDataSet(sqlCmd)
                        If dsTax IsNot Nothing AndAlso dsTax.Tables.Count > 0 AndAlso dsTax.Tables(0).Rows.Count > 0 Then
                            cgstPercent = Convert.ToString(dsTax.Tables(0).Rows(0)("CGST_PERCENT"))
                            sgstPercent = Convert.ToString(dsTax.Tables(0).Rows(0)("SGST_PERCENT"))
                            igstPercent = Convert.ToString(dsTax.Tables(0).Rows(0)("IGST_PERCENT"))
                            utgstPercent = Convert.ToString(dsTax.Tables(0).Rows(0)("UTGST_PERCENT"))
                            CompensationCessPercent = Convert.ToString(dsTax.Tables(0).Rows(0)("COMPENSATION_CESS_PERCENT"))
                            dsTax.Dispose()
                        End If
                        sqlCmd.Dispose()
                        mstrItemText = mstrItemText & StrItemCode & "|" & strDrawingNo & "|" & strDescription & "|" & dblQuantity & "|" & dbldespatchQty & "|" & strCustRef & "|" & StrAmendmentNo & "|" & strSchDate & "|" & strSchTime & "|" & strAccountCode & "|" & strSrvDINo & "|" & hsnsacType & "|" & hsnsacCode & "|" & cgstType & "|" & cgstPercent & "|" & sgstType & "|" & sgstPercent & "|" & igstType & "|" & igstPercent & "|" & utgstType & "|" & utgstPercent & "|" & CompensationCessType & "|" & CompensationCessPercent & "^"
                    Else
                        mstrItemText = mstrItemText & StrItemCode & "|" & strDrawingNo & "|" & strDescription & "|" & dblQuantity & "|" & dbldespatchQty & "|" & strCustRef & "|" & StrAmendmentNo & "|" & strSchDate & "|" & strSchTime & "|" & strAccountCode & "|" & strSrvDINo & "^"
                    End If
                    '101188073 End
                End If
            Next intSubItem
        End With
        If Len(mstrItemText) = 0 Then
            Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Me.SpItems.Focus()
            Exit Sub
        End If
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub frmMKTTRN0035a_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        SpItems.CursorType = FPSpreadADO.CursorTypeConstants.CursorTypeColHeader
        SpItems.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        SpItems.Focus()
    End Sub

    Private Sub frmMKTTRN0035a_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim intCount As Short
        'Added By ekta uniyal on 4 Mar 2014
        If ((CBool(Find_Value("SELECT ISNULL(ALLOWBACKDATE_TOYOTA_PDS,0)as ALLOWBACKDATE_TOYOTA_PDS FROM SALES_PARAMETER where unit_code = '" & gstrUNITID & "'")) = True) And (Val(Find_Value("SELECT ISNULL(NOOFBACKDAYS_TOYOTA_PDS,0) AS  NOOFBACKDAYS_TOYOTA_PDS FROM SALES_PARAMETER where unit_code = '" & gstrUNITID & "'")) > 0)) Then
            mblnBackDatePDS = True
            intCount = Val(Find_Value("SELECT ISNULL(NOOFBACKDAYS_TOYOTA_PDS,0) AS  NOOFBACKDAYS_TOYOTA_PDS FROM SALES_PARAMETER  where unit_code = '" & gstrUNITID & "'"))
            'End Here
            mstartingdate = DateAdd(DateInterval.Day, -intCount, GetServerDate())
        Else
            mblnBackDatePDS = False
            mstartingdate = GetServerDate()
        End If
        Call AddColumnsInSpread()
        'Grpselectunselect.Enabled = False
        'optunselectall.Checked = True
        TxtPDSNo.Enabled = False : TxtPDSNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtSONo.Enabled = False : TxtSONo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 4.4)
        mstrItemText = ""
        '101188073 Start
        _itemCode = String.Empty
        '101188073 End
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Sub AddColumnsInSpread()
        With SpItems
            .MaxRows = 0
            If gblnGSTUnit Then  '101188073 Start
                .MaxCols = GridHeader.COMPENSATION_CESS_TYPE
            Else
                .MaxCols = GridHeader.FTSBARCODE
            End If  '101188073 End
            .UserColAction = FPSpreadADO.UserColActionConstants.UserColActionSort
            .Row = 0
            .Col = GridHeader.Mark : .Text = "Mark" : .set_ColWidth(GridHeader.Mark, 4)
            .Col = GridHeader.ItemCode : .Text = "Item Code" : .set_ColWidth(GridHeader.ItemCode, 10)
            .Col = GridHeader.DrawingNo : .Text = "Drawing No" : .set_ColWidth(GridHeader.DrawingNo, 12)
            .Col = GridHeader.Description : .Text = "Description" : .set_ColWidth(GridHeader.Description, 16)
            .Col = GridHeader.Quantity : .Text = "PDS Qty" : .set_ColWidth(GridHeader.Quantity, 6)
            .Col = GridHeader.DespatchQty : .Text = "Disp Qty" : .set_ColWidth(GridHeader.DespatchQty, 6)
            .Col = GridHeader.CustRef : .Text = "Cust Ref" : .set_ColWidth(GridHeader.CustRef, 10)
            .Col = GridHeader.AmendmentNo : .Text = "Amendment No." : .set_ColWidth(GridHeader.AmendmentNo, 8)
            .Col = GridHeader.SchDate : .Text = "Sch Date" : .set_ColWidth(GridHeader.SchDate, 8)
            .Col = GridHeader.SChTime : .Text = "Sch Time" : .set_ColWidth(GridHeader.SChTime, 6)
            .Col = GridHeader.AccountCode : .Text = "Account Code" : .set_ColWidth(GridHeader.AccountCode, 12)
            .Col = GridHeader.FTSITEM : .Text = "FTS ITEM" : .set_ColWidth(GridHeader.FTSITEM, 4)
            .Col = GridHeader.FTSBARCODE : .Text = "FTS BarCode" : .set_ColWidth(GridHeader.FTSBARCODE, 4)
            '101188073 Start
            If gblnGSTUnit Then
                .Col = GridHeader.IS_HSN_SAC : .Text = "HSN/SAC" : .set_ColWidth(GridHeader.IS_HSN_SAC, 8)
                .Col = GridHeader.HSN_SAC_CODE : .Text = "HSN/SAC No." : .set_ColWidth(GridHeader.HSN_SAC_CODE, 12)
                .Col = GridHeader.CGST_TYPE : .Text = "CGST" : .set_ColWidth(GridHeader.CGST_TYPE, 8)
                .Col = GridHeader.SGST_TYPE : .Text = "SGST" : .set_ColWidth(GridHeader.SGST_TYPE, 8)
                .Col = GridHeader.IGST_TYPE : .Text = "IGST" : .set_ColWidth(GridHeader.IGST_TYPE, 8)
                .Col = GridHeader.UTGST_TYPE : .Text = "UTGST" : .set_ColWidth(GridHeader.UTGST_TYPE, 8)
                .Col = GridHeader.COMPENSATION_CESS_TYPE : .Text = "Comp. CESS" : .set_ColWidth(GridHeader.COMPENSATION_CESS_TYPE, 8)
            End If
            '101188073 End
            Grpselectunselect.Enabled = False
            Me.optunselectall.Checked = True
        End With
    End Sub

    Public Function SelectDatafromItem_Mst(Optional ByRef pstrItem As String = "", Optional ByRef intAlreadyItem As Integer = 0) As Object
        On Error GoTo ErrHandler
        Dim strItembal As String
        Dim rsItembal As ClsResultSetDB
        Dim INTCOUNT As Short
        Dim intRecordCount As Integer 'To Hold Record Count

        'Added By Ekta Uniyal on 4 Mar 2014
        strItembal = "select Distinct PDSNO, m.item_code, m.cust_drgNo, m.Description, m.Quantity, m.Cust_Ref, m.amendment_no,( CONVERT( VARCHAR(14),PDSDATE,106)+' '+  CONVERT(VARCHAR(6),SCH_TIME))  AS 'pdsdate',"
        strItembal = strItembal & " case when Sch_Time = '23:59' then '' else Sch_Time end as sch_time, m.UNLOC, m.USLOC, m.Account_code "
        strItembal = strItembal & " ,despatch_qty=  (select isnull(sum(sd.sales_quantity),0) from sales_dtl sd  ,saleschallan_dtl sc where sc.Unit_code=sd.Unit_code and sc.Unit_code='" & gstrUNITID & "' and sc.doc_no=sd.doc_no and sd.srvdino=pdsno and sc.cancel_flag=0 and item_code=m.item_code and m.cust_drgno=cust_drgno)"
        strItembal = strItembal & " , fts_item,FTS_BARCODE_TRACKING "
        '101188073 Start
        If gblnGSTUnit Then
            strItembal = strItembal & " ,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS"
        End If
        '101188073 End
        strItembal = strItembal & " from vw_ToyotaPDS_Help_HILEX m (NOLOCK) "
        If mblnMultipleSOPDS = True Then
            strItembal = strItembal & " where m.Unit_code = '" & gstrUNITID & "' and  ( M.PDSNO ='" & Me.TxtPDSNo.Text & "'  AND  M.PDSDATE >= '" & getDateForDB(mstartingdate) & "'  and m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.unit_code = b.unit_code and  a.location_code = b.location_code and a.doc_no=b.doc_no where a.unit_code = '" & gstrUNITID & "' and  m.pdsno = b.srvdino and a.cancel_flag <> 1  and b.item_code=m.item_code and b.cust_item_code=m.cust_drgno and a.account_code=m.account_code)))"
        Else
            strItembal = strItembal & " where m.Unit_code = '" & gstrUNITID & "' and  ( M.PDSNO ='" & Me.TxtPDSNo.Text & "' AND M.CUST_REF='" & Me.TxtSONo.Text & " '  AND  M.PDSDATE >= CONVERT(CHAR(13), '" & mstartingdate & "' , 106) ) and m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.unit_code = b.unit_code and a.location_code = b.location_code and a.doc_no=b.doc_no where a.unit_code = '" & gstrUNITID & "' and m.pdsno = b.srvdino and a.cancel_flag <> 1  and b.item_code=m.item_code and b.cust_item_code=m.cust_drgno and a.account_code=m.account_code))"
        End If

        strItembal = strItembal & " order by pdsdate desc"
        'End Here
        rsItembal = New ClsResultSetDB
        If Len(Trim(strItembal)) <= 0 Then Exit Function
        mP_Connection.CommandTimeout = 0
        rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        mP_Connection.CommandTimeout = 30
        intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rsItembal.MoveFirst() 'move to first record
            For INTCOUNT = 1 To intRecordCount
                With SpItems
                    .MaxRows = .MaxRows + 1
                    .Row = INTCOUNT
                    .Col = GridHeader.Mark : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
                    .Col = GridHeader.ItemCode
                    .Text = rsItembal.GetValue("item_code") : .Lock = True
                    .Col = GridHeader.DrawingNo
                    .Text = rsItembal.GetValue("cust_drgNo") : .Lock = True
                    .Col = GridHeader.Description
                    .Text = rsItembal.GetValue("Description") : .Lock = True
                    .Col = GridHeader.Quantity
                    .Text = rsItembal.GetValue("Quantity") : .Lock = True
                    .Col = GridHeader.DespatchQty
                    .Text = rsItembal.GetValue("despatch_qty") : .Lock = True
                    .Col = GridHeader.CustRef
                    .Text = rsItembal.GetValue("Cust_Ref") : .Lock = True
                    .Col = GridHeader.AmendmentNo
                    .Text = rsItembal.GetValue("amendment_no") : .Lock = True
                    .Col = GridHeader.SchDate
                    .Text = rsItembal.GetValue("PDSDATE") : .Lock = True
                    .Col = GridHeader.AccountCode
                    .Text = rsItembal.GetValue("Account_Code") : .Lock = True
                    .Col = GridHeader.FTSITEM
                    .Text = rsItembal.GetValue("fts_item") : .Lock = True
                    .Col = GridHeader.FTSBARCODE
                    .Text = rsItembal.GetValue("FTS_BARCODE_TRACKING") : .Lock = True
                    If rsItembal.GetValue("fts_item") = True And rsItembal.GetValue("FTS_BARCODE_TRACKING") = True Then
                        .BlockMode = True : .Row = INTCOUNT : .Row2 = INTCOUNT : .Col = 1 : .Col2 = .MaxCols : .BackColor = System.Drawing.Color.Bisque : .BlockMode = False
                    ElseIf rsItembal.GetValue("fts_item") = True And rsItembal.GetValue("FTS_BARCODE_TRACKING") = False Then
                        .BlockMode = True : .Row = INTCOUNT : .Row2 = INTCOUNT : .Col = 1 : .Col2 = .MaxCols : .BackColor = System.Drawing.Color.LightSkyBlue : .BlockMode = False
                    End If
                    '101188073 Start
                    If gblnGSTUnit Then
                        .Row = INTCOUNT
                        .Col = GridHeader.IS_HSN_SAC
                        .Text = rsItembal.GetValue("ISHSNORSAC") : .Lock = True
                        .Col = GridHeader.HSN_SAC_CODE
                        .Text = rsItembal.GetValue("HSNSACCODE") : .Lock = True
                        .Col = GridHeader.CGST_TYPE
                        .Text = rsItembal.GetValue("CGSTTXRT_TYPE") : .Lock = True
                        .Col = GridHeader.SGST_TYPE
                        .Text = rsItembal.GetValue("SGSTTXRT_TYPE") : .Lock = True
                        .Col = GridHeader.IGST_TYPE
                        .Text = rsItembal.GetValue("IGSTTXRT_TYPE") : .Lock = True
                        .Col = GridHeader.UTGST_TYPE
                        .Text = rsItembal.GetValue("UTGSTTXRT_TYPE") : .Lock = True
                        .Col = GridHeader.COMPENSATION_CESS_TYPE
                        .Text = rsItembal.GetValue("COMPENSATION_CESS") : .Lock = True
                    End If
                    '101188073 End
                End With
                rsItembal.MoveNext()
            Next INTCOUNT
            rsItembal.ResultSetClose()
            rsItembal = Nothing
        Else
            MsgBox("No Records Found", MsgBoxStyle.Information, "eMPro")
            Exit Function
        End If
        'If Len(pstrItem) > 0 Then Call SelectPreviousItem(pstrItem)

        SelectDatafromItem_Mst = mstrItemText
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

    Private Sub optCustDrawNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCustDrawNo.CheckedChanged
        If eventSender.Checked Then
            If optCustDrawNo.Checked = True Then
                With SpItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader.DrawingNo)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If

            txtsearch.Text = ""
            txtsearch.Focus()

        End If
    End Sub
    Private Sub optCustRef_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCustRef.CheckedChanged
        If eventSender.Checked Then
            txtsearch.Text = ""
            txtsearch.Focus()
        End If
    End Sub

    Private Sub optDescription_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDescription.CheckedChanged
        If eventSender.Checked Then
            txtsearch.Text = ""
            txtsearch.Focus()

        End If
    End Sub

    Private Sub OptItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItemCode.CheckedChanged
        If eventSender.Checked Then
            If optItemCode.Checked = True Then
                With SpItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader.ItemCode)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If

            txtsearch.Text = ""
            txtsearch.Focus()

        End If
    End Sub

    Private Sub optKanbanNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If eventSender.Checked Then
            txtsearch.Text = ""
        End If
    End Sub

    Private Sub optSchDate_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSchDate.CheckedChanged
        If eventSender.Checked Then
            txtsearch.Text = ""
            txtsearch.Focus()
        End If
    End Sub

    Private Sub SpItems_Change(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpItems.Change
        Call SpItems_ClickEvent(SpItems, New AxFPSpreadADO._DSpreadEvents_ClickEvent(GridHeader.Mark, eventArgs.row))
    End Sub

    Private Sub SpItems_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SpItems.ClickEvent
        Dim varItemCode As Object
        Dim varCustref As Object
        Dim blnflag As Boolean
        Dim blnflag1 As Boolean
        Dim intSubItem As Short
        Dim varToolCost As Object

        Dim intToolCost As Short
        Dim intOtherItemToolCost As Short
        Dim intSubItem1 As Short
        Dim blnOtherRecord As Boolean
        On Error GoTo ErrHandler

        With SpItems
            If eventArgs.col = GridHeader.Mark Then
                .Row = eventArgs.row : .Col = eventArgs.col
                If CBool(.Value) = False Then Exit Sub
                varItemCode = Nothing
                blnflag = .GetText(GridHeader.ItemCode, eventArgs.row, varItemCode)

                For intSubItem = 1 To .MaxRows
                    .Row = intSubItem
                    .Col = GridHeader.Mark
                    If CBool(.Value) = True And .Row <> eventArgs.row Then
                        '.Col = GridHeader.ItemCode


                        'If UCase(Trim(varItemCode)) = UCase(Trim(.Text)) Then
                        .Col = GridHeader.CustRef
                        varCustref = Nothing
                        blnflag1 = .GetText(GridHeader.CustRef, eventArgs.row, varCustref)

                        If mblnMultipleSOPDS = False And UCase(Trim(varCustref)) <> UCase(Trim(.Text)) Then
                            MsgBox("You can not select different Sales Orders. ", MsgBoxStyle.Information, "eMPro")
                            .Col = GridHeader.Mark
                            .Row = eventArgs.row
                            .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                            System.Windows.Forms.Application.DoEvents()
                            'End If
                            Exit Sub
                        End If


                    End If
                Next
            End If
        End With
        Exit Sub

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub SpItems_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpItems.KeyPressEvent
        With SpItems
            If e.keyAscii = 18 Or e.keyAscii = System.Windows.Forms.Keys.Space Then
                .Col = 1
                .Value = IIf(Val(.Value), False, True)
                Call SpItems_ClickEvent(SpItems, New AxFPSpreadADO._DSpreadEvents_ClickEvent(GridHeader.Mark, .Row))
            End If
        End With
    End Sub

    Private Sub SpItems_KeyUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpItems.KeyUpEvent
        With SpItems
            If eventArgs.keyCode = 18 Or eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
                .Col = 1
                .Value = IIf(Val(.Value), False, True)
                Call SpItems_ClickEvent(SpItems, New AxFPSpreadADO._DSpreadEvents_ClickEvent(GridHeader.Mark, .Row))
            End If
        End With
    End Sub

    Private Sub SpItems_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpItems.LeaveCell
        With SpItems
            .Row = -1
            .Col = -1
            '  .BackColor = System.Drawing.Color.White
            ' .ForeColor = System.Drawing.Color.Black
            .Col = -1
            .Row = IIf(eventArgs.newRow <= 0, 1, eventArgs.newRow)
            '.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000D)
            '.ForeColor = System.Drawing.Color.White
        End With
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
    End Sub

    Sub SearchItem()
        On Error GoTo ErrHandler
        Dim intCount As Short
        With SpItems
            .Row = -1
            .Col = -1
            .BackColor = System.Drawing.Color.White
            .ForeColor = System.Drawing.Color.Black

            If optItemCode.Checked Then
                .Col = 2
            End If
            If optDescription.Checked Then
                .Col = 4
            End If
            If optCustRef.Checked Then
                .Col = 7
            End If
            If optSchDate.Checked Then
                .Col = 9
            End If
            If optCustDrawNo.Checked Then
                .Col = 3
            End If

            For intCount = 1 To .MaxRows
                .Row = intCount
                If UCase(Mid(.Text, 1, Len(txtsearch.Text))) = UCase(txtsearch.Text) Then
                    .TopRow = .Row
                    .Col = -1
                    .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000D)
                    .ForeColor = System.Drawing.Color.White
                    Exit Sub
                End If
            Next
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred

    End Sub

    Private Sub txtSearch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtsearch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtsearch_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtsearch.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        With SpItems
            If KeyCode = 13 And Len(Trim(txtsearch.Text)) > 0 Then
                .Col = 1
                .Value = IIf(CBool(.Value), False, True)
            End If
        End With
    End Sub

    Function ValidateData() As Boolean
        Dim intSubItem As Short
        Dim blnOtherRecord As Boolean
        Dim strMessage As String

        Dim strAccountCode As String
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim strSalesTaxType As String
        Dim strExciseType As String
        Dim StrItemCode As String
        Dim strSOType As String

        Dim strOtherAccountCode As String
        Dim strOtherCustRef As String
        Dim strOtherAmendmentNo As String
        Dim strOtherSalesTaxType As String
        Dim strOtherExciseType As String
        Dim strOtherItemCode As String
        Dim strOtherSOType As String
        Dim strftsitem As String
        Dim strftsbarcode As String
        Dim strotherftsitem As String
        Dim strotherftsbarcode As String

        With SpItems
            For intSubItem = 1 To .MaxRows
                .Row = intSubItem
                .Col = GridHeader.Mark
                If CBool(.Value) = True Then
                    If Not blnOtherRecord Then
                        .Col = GridHeader.AccountCode
                        strAccountCode = Trim(.Text)
                        .Col = GridHeader.CustRef
                        strCustRef = Trim(.Text)
                        .Col = GridHeader.AmendmentNo
                        StrAmendmentNo = Trim(.Text)
                        .Col = GridHeader.ItemCode
                        StrItemCode = Trim(.Text)
                        'Added By ekta uniyal on 4 Mar 2014
                        strSalesTaxType = Trim(Find_Value("select SalesTax_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and amendment_no='" & StrAmendmentNo & "'"))
                        strSOType = Trim(Find_Value("select PO_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and amendment_no='" & StrAmendmentNo & "'"))
                        strExciseType = Trim(Find_Value("select Excise_duty from cust_ord_dtl where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and Amendment_no ='" & StrAmendmentNo & "' and item_code ='" & StrItemCode & "' and cust_drgNo='" & StrItemCode & "'"))
                        .Col = GridHeader.FTSITEM
                        strftsitem = Trim(.Text)
                        .Col = GridHeader.FTSBARCODE
                        strftsbarcode = Trim(.Text)
                        'End Here
                        blnOtherRecord = True
                    Else
                        .Col = GridHeader.AccountCode
                        strOtherAccountCode = Trim(.Text)
                        .Col = GridHeader.CustRef
                        strOtherCustRef = Trim(.Text)
                        .Col = GridHeader.AmendmentNo
                        strOtherAmendmentNo = Trim(.Text)
                        .Col = GridHeader.ItemCode
                        strOtherItemCode = Trim(.Text)
                        'Added By ekta uniyal on 4 Mar 2014
                        strOtherSalesTaxType = Trim(Find_Value("select SalesTax_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and amendment_no='" & strOtherAmendmentNo & "'"))
                        strOtherSOType = Trim(Find_Value("select PO_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and amendment_no='" & strOtherAmendmentNo & "'"))
                        strOtherExciseType = Trim(Find_Value("select Excise_duty from cust_ord_dtl where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and Amendment_no ='" & strOtherAmendmentNo & "' and item_code ='" & strOtherItemCode & "' and cust_drgNo='" & strOtherItemCode & "'"))
                        .Col = GridHeader.FTSITEM
                        strotherftsitem = Trim(.Text)
                        .Col = GridHeader.FTSBARCODE
                        strotherftsbarcode = Trim(.Text)
                        'End Here

                        If UCase(strAccountCode) <> UCase(strOtherAccountCode) Then
                            strMessage = "Two or more SOs of different Customers are not allowed." & vbCrLf
                            strMessage = strMessage & "1. " & strAccountCode & " -> " & strCustRef & vbCrLf
                            strMessage = strMessage & "2. " & strOtherAccountCode & " -> " & strOtherCustRef
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If

                        If UCase(strSOType) <> UCase(strOtherSOType) Then
                            strMessage = "OEM(O) and Spare(S) type SOs can not be included in same invoice." & vbCrLf
                            strMessage = strMessage & "1. " & strCustRef & " -> " & strSOType & vbCrLf
                            strMessage = strMessage & "2. " & strOtherCustRef & " -> " & strOtherSOType
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If
                        If Not gblnGSTUnit Then   '101188073
                            If UCase(strSalesTaxType) <> UCase(strOtherSalesTaxType) Then
                                strMessage = "Two or more SOs can not have different Sales Tax Rate." & vbCrLf
                                strMessage = strMessage & "1. " & strCustRef & " -> " & strSalesTaxType & vbCrLf
                                strMessage = strMessage & "2. " & strOtherCustRef & " -> " & strOtherSalesTaxType
                                MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                                ValidateData = False
                                Exit Function
                            End If

                            If UCase(strExciseType) <> UCase(strOtherExciseType) Then
                                strMessage = "Two or more SOs can not have different Excise Rate." & vbCrLf
                                strMessage = strMessage & "1. " & strCustRef & " -> " & strExciseType & vbCrLf
                                strMessage = strMessage & "2. " & strOtherCustRef & " -> " & strOtherExciseType
                                MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                                ValidateData = False
                                Exit Function
                            End If '101188073
                        End If
                        If frmMKTTRN0071_HILEX.OptNormalDispatch.Checked = True Then
                            If (strftsitem <> strotherftsitem) Or (strftsbarcode <> strotherftsbarcode) Then
                                strMessage = "Two or more Combination  can not have in one invoice." & vbCrLf
                                MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                                ValidateData = False
                                Exit Function
                            End If
                        End If


                        End If
                End If
            Next intSubItem
            ValidateData = True
        End With
    End Function

    Public Function Find_Value(ByRef strField As String) As String

        On Error GoTo ErrHandler
        Dim rs As New ADODB.Recordset
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If rs.RecordCount > 0 Then
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            If IsDBNull(rs.Fields(0).Value) = False Then
                Find_Value = rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function


    Private Sub CmdPDSNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdPDSNo.Click
        On Error GoTo ErrHandler
        Dim strHelp As String
        Dim rssalechallan As ClsResultSetDB
        Dim salechallan As String
        Dim strInvoiceType As Object

        If Len(Me.TxtPDSNo.Text) = 0 Then
            'Added By ekta uniyal on 4 Mar 2014
            strHelp = ShowList(1, (TxtPDSNo.MaxLength), "", "PDSNo ", "PDSDATE", "vw_ToyotaPDS_Help_HILEX M", "AND ( M.PDSDATE >= CONVERT(CHAR(13), '" & mstartingdate.ToString("dd MMM yyyy") & "' , 106) ) and m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.Unit_code = b.unit_code and a.location_code = b.location_code and a.doc_no=b.doc_no where a.unit_code = '" & gstrUNITID & "' and  m.PDSNO = b.srvdino and a.cancel_flag <> 1 and m.item_code=b.item_code and m.cust_drgno=b.Cust_Item_Code and m.account_code=a.account_code  ) )")
            'End Here

            If strHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Exit Sub
            Else
                AddColumnsInSpread()
                TxtSONo.Text = ""
                TxtPDSNo.Text = strHelp
                mblnMultipleSOPDS = False
                If ((CBool(Find_Value("SELECT ISNULL(MULTIPLE_SO_PDS_TOYOTA,0)as MULTIPLE_SO_PDS_TOYOTA FROM SALES_PARAMETER where unit_code = '" & gstrUNITID & "'")) = True)) Then
                    If ((CBool(Find_Value("SELECT ISNULL(PDS_TOYOTA_CUSTOMER,0)as PDS_TOYOTA_CUSTOMER  FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUNITID & "'AND CUSTOMER_CODE IN(SELECT TOP 1 ACCOUNT_CODE FROM vw_ToyotaPDS_Help_HILEX WHERE UNIT_CODE='" & gstrUNITID & "' AND PDSNo='" & strHelp & "')")) = True)) Then
                        mblnMultipleSOPDS = True
                    End If
                End If
                If ((CBool(Find_Value("SELECT ISNULL(PDSNagaare,0)as PDSNagaare FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUNITID & "'AND CUSTOMER_CODE IN(SELECT TOP 1 ACCOUNT_CODE FROM vw_ToyotaPDS_Help_HILEX WHERE UNIT_CODE='" & gstrUNITID & "' AND PDSNo='" & strHelp & "')")) = True)) Then
                    mblnMultipleSOPDS = True
                End If
                If mblnMultipleSOPDS = True Then
                    CmdSONO.Enabled = False : TxtSONo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Else
                    CmdSONO.Enabled = True : TxtSONo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
            End If

        Else
            'Added by ekta uniyal on 4 mar 2014
            strHelp = ShowList(1, (TxtPDSNo.MaxLength), TxtPDSNo.Text, "PDSNo", "PDSDATE", "vw_ToyotaPDS_Help_HILEX M ", "AND ( M.PDSDATE >= CONVERT(CHAR(13), '" & mstartingdate.ToString("dd MMM yyyy") & "' , 106) ) and m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.Unit_code = b.unit_code and a.location_code = b.location_code and a.doc_no=b.doc_no where a.unit_code = '" & gstrUNITID & "' and m.PDSNO = b.srvdino and a.cancel_flag <> 1 ) + (select IsNull(sum(sales_quantity),0) as sales_quantity  from printedsrv_dtl p where p.UNIT_CODE = '" & gstrUNITID & "' AND p.KanBan_No=m.PDSNO)+(Select isnull(Sum(quantity),0) as sales_quantity From mkt_57F4challankanban_dtl B inner join mkt_57F4challan_hdr A on A.Unit_code = B.unit_code and B.doc_type=A.doc_type and B.doc_no = A.doc_no where A.unit_code = '" & gstrUNITID & "' and A.cancel_flag = 0 and B.Kanban_no=m.PDSNO and M.PDSNO='" & Me.TxtPDSNo.Text & "'))")
            'End Here
            If strHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Exit Sub
            Else
                AddColumnsInSpread()
                TxtSONo.Text = ""
                TxtPDSNo.Text = strHelp
                'Call SelectDatafromItem_Mst()
            End If
        End If
        If mblnMultipleSOPDS = True Then
            If Len(Me.TxtPDSNo.Text) > 0 Then
                Call SelectDatafromItem_Mst()
                Grpselectunselect.Enabled = True
            Else
                MsgBox("Select PDS No First ", MsgBoxStyle.Information, "eMPro")
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdSONO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSONO.Click

        On Error GoTo ErrHandler
        Dim strHelp As String
        Dim rssalechallan As ClsResultSetDB
        Dim salechallan As String
        Dim strInvoiceType As Object
        ' Issue ID 10119746 starts

        TxtSONo.Text = ""
        AddColumnsInSpread()
        If Len(Me.TxtPDSNo.Text) > 0 Then
            If Len(Me.TxtSONo.Text) = 0 Then
                strHelp = ShowList(1, (TxtPDSNo.MaxLength), "", "CUST_REF ", "AMENDMENT_NO", "vw_ToyotaPDS_Help_HILEX M", "AND ( M.PDSNO = '" & Me.TxtPDSNo.Text & "' )")
                If strHelp = "-1" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Exit Sub
                Else
                    TxtSONo.Text = strHelp
                    Call SelectDatafromItem_Mst()
                    Grpselectunselect.Enabled = True
                End If
            Else
                strHelp = ShowList(1, (TxtPDSNo.MaxLength), TxtPDSNo.Text, "CUST_REF", "AMENDMENT_NO", "vw_ToyotaPDS_Help_HILEX M ", "AND ( M.PDSNO = '" & Me.TxtPDSNo.Text & "'  and m.cust_ref='" & Me.TxtSONo.Text & "' ) ")
                If strHelp = "-1" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    TxtSONo.Text = ""
                    Exit Sub
                Else
                    TxtSONo.Text = strHelp
                    Call SelectDatafromItem_Mst()
                    Grpselectunselect.Enabled = True
                End If
            End If
        Else
            MsgBox("Select PDS No First ", MsgBoxStyle.Information, "eMPro")
        End If
        Exit Sub
ErrHandler:  ' Issue ID 10119746 end 
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub


    Private Sub optselectall_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optselectall.CheckedChanged
        ' Issue ID 10119746 starts
        If sender.Checked Then
            On Error GoTo ErrHandler
            If optselectall.Checked = True Then
                With SpItems
                    .BlockMode = True
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 1
                    .Text = CStr(System.Windows.Forms.CheckState.Checked)
                    .BlockMode = False
                End With
            End If
            Exit Sub
            ' Issue ID 10119746 end 
ErrHandler: Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
            Exit Sub
        End If

    End Sub

    Private Sub optunselectall_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optunselectall.CheckedChanged
        ' Issue ID 10119746 starts
        If sender.Checked Then
            On Error GoTo ErrHandler
            If optunselectall.Checked = True Then
                With SpItems
                    .BlockMode = True
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 1
                    .Text = CStr(System.Windows.Forms.CheckState.Unchecked)
                    .BlockMode = False
                End With
            End If
            Exit Sub
            ' Issue ID 10119746 end 
ErrHandler: Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
            Exit Sub
        End If

    End Sub
    '101188073 Start
    Private Function ValidateGSTTaxes() As Boolean
        Dim blnResult As Boolean = True
        Dim dtGST As New DataTable
        Dim dtGSTHSN As New DataTable
        Dim drGST As DataRow
        Dim drGSTHSN As DataRow
        Dim strItems As String()
        Try
            With SpItems
                dtGST.Columns.Add("ITEM_CODE", GetType(String))
                dtGST.Columns.Add("UNIT_CODE", GetType(String))
                dtGSTHSN.Columns.Add("ITEM_CODE", GetType(String))
                dtGSTHSN.Columns.Add("UNIT_CODE", GetType(String))
                For rowIndex As Integer = 1 To .MaxRows
                    .Row = rowIndex
                    .Col = GridHeader.Mark
                    If CBool(.Value) Then
                        drGST = dtGST.NewRow()
                        drGSTHSN = dtGSTHSN.NewRow()
                        .Row = rowIndex
                        .Col = GridHeader.ItemCode
                        drGST("ITEM_CODE") = .Text
                        drGST("UNIT_CODE") = gstrUnitId
                        drGSTHSN("ITEM_CODE") = .Text
                        drGSTHSN("UNIT_CODE") = gstrUnitId
                        dtGST.Rows.Add(drGST)
                        .Row = rowIndex
                        .Col = GridHeader.HSN_SAC_CODE
                        If Len(Trim(.Text)) = 0 Then
                            dtGSTHSN.Rows.Add(drGSTHSN)
                        End If
                    End If
                Next
                If dtGST IsNot Nothing AndAlso dtGST.Rows.Count > 0 Then
                    If Not ValidateGSTItemWise(dtGST, dtGSTHSN) Then
                        If Len(_itemCode) > 0 Then
                            strItems = _itemCode.Split(",")
                            If strItems IsNot Nothing AndAlso strItems.Count > 0 Then
                                If dtGST.Rows.Count = strItems.Count Then
                                    blnResult = False
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            blnResult = False
        Finally
            dtGST.Dispose()
            dtGSTHSN.Dispose()
        End Try
        Return blnResult
    End Function
    Private Function ValidateGSTItemWise(ByRef dtGST As DataTable, ByRef dtGSTHSN As DataTable) As Boolean
        Dim blnResult As Boolean = True
        Dim sqlCmd As SqlCommand
        _itemCode = String.Empty
        Try
            sqlCmd = New SqlCommand()
            sqlCmd.CommandType = CommandType.StoredProcedure
            sqlCmd.CommandText = "USP_VALIDATE_ITEM_GST"
            sqlCmd.Parameters.AddWithValue("@ITEM_CODE", dtGST)
            sqlCmd.Parameters.AddWithValue("@ITEM_NOT_HSN", dtGSTHSN)
            sqlCmd.Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Value = String.Empty
            sqlCmd.Parameters("@MESSAGE").Direction = ParameterDirection.InputOutput
            sqlCmd.Parameters.Add("@ITEMS", SqlDbType.VarChar, 8000).Value = String.Empty
            sqlCmd.Parameters("@ITEMS").Direction = ParameterDirection.InputOutput
            SqlConnectionclass.ExecuteNonQuery(sqlCmd)
            If Len(sqlCmd.Parameters("@MESSAGE").Value) > 0 AndAlso Len(sqlCmd.Parameters("@ITEMS").Value) > 0 Then
                _itemCode = sqlCmd.Parameters("@ITEMS").Value
                MsgBox(sqlCmd.Parameters("@MESSAGE").Value.ToString(), MsgBoxStyle.Information, "eMPro")
                blnResult = False
            End If
        Catch ex As Exception
            blnResult = False
        End Try
        Return blnResult
    End Function
    '101188073 End
End Class