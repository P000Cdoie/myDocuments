Imports System.Data
Imports System.Data.SqlClient
'Modified by Nitin Mehta on 19/Dec/2011 modified for CHANGE MANAGEMENT
' Modified By Deepak Kumar on 31 jan 2012 to support multiunit change management
'---------------------------------------------------------------------------------------
'REVISED BY :   SHUBHRA 
'REVISED ON :   09 SEP 2013
'ISSUE ID   :   10449968
'DESCRIPTION:   Error while searching Sch Date in Item Help of Invoice against Nagare
'---------------------------------------------------------------------------------------
''Revised By:       Saurav Kumar
''Revised On:       04 Oct 2013
''Issue ID  :       10462231 - eMpro ISuite Changes
'***********************************************************************************************************************************
''Revised By    :       Geetanjali Aggrawal
''Revised On    :       03 Mar 2014 
''Purpose       :       for HILEX migration

'REVISED BY     :  ASHISH SHARMA
'REVISED DATE   :  29 MAY 2017
'ISSUE ID       :  101188073 
'PURPOSE        :  GST changes
'***********************************************************************************************************************************
Public Class FrmHelpNagare
    Dim mStrSql As String = ""
    Dim mRowIndex As Integer = -1
    Dim dt As DataTable
    Dim dtSelectedRows As DataTable
    Dim mblnFtsSpareDispatch As Boolean
    Dim mblnFtsEnabled As Boolean
    '101188073 Start
    Dim _itemCode As String
    '101188073 End

    Public Property FTSSpareDispatch() As Boolean
        Get
            FTSSPAREDISPATCH = mblnFtsSpareDispatch
        End Get
        Set(ByVal Value As Boolean)
            mblnFtsSpareDispatch = Value
        End Set
    End Property
    Public Property FTSEnabled() As Boolean
        Get
            FTSEnabled = mblnFtsEnabled
        End Get
        Set(ByVal Value As Boolean)
            mblnFtsEnabled = Value
        End Set
    End Property

    Enum GridHeader
        Mark = 0
        KanbanNo = 1
        ItemCode = 2
        DrawingNo = 3
        Description = 4
        Quantity = 5
        CustRef = 6
        AmendmentNo = 7
        SchDate = 8
        SChTime = 9
        UNLoc = 10
        USLOC = 11
        AccountCode = 12
        Tool_Cost = 13
        FTS_ITEM = 14
        FTS_BARCODE = 15
        '101188073 Start
        IS_HSN_SAC = 16
        HSN_SAC_CODE = 17
        CGST_TYPE = 18
        SGST_TYPE = 19
        IGST_TYPE = 20
        UTGST_TYPE = 21
        COMPENSATION_CESS_TYPE = 22
        BATCH_CODE = 23
        '101188073 End
    End Enum

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        dt.DefaultView.RowFilter = "" & Me.cboSearchBy.Text & " like '" & Me.txtSearch.Text.Trim & "%'"
        Me.GrdHelp.DataSource = dt
    End Sub
    Private Sub FillHelp()
        Dim i As Integer
        Try
            dt = SqlConnectionclass.GetDataTable(mStrSql)
            Me.GrdHelp.DataSource = dt

            Dim column As DataGridColumnStyle
            GrdHelp.Columns(0).Width = 25
            GrdHelp.Columns(1).Width = 110
            GrdHelp.Columns(2).Width = 60
            GrdHelp.Columns(3).Width = 110
            GrdHelp.Columns(4).Width = 110
            GrdHelp.Columns(5).Width = 110
            GrdHelp.Columns(6).Width = 45
            GrdHelp.Columns(7).Width = 70
            GrdHelp.Columns(8).Width = 40
            GrdHelp.Columns(9).Width = 55
            GrdHelp.Columns(10).Width = 50
            GrdHelp.Columns(11).Width = 40
            GrdHelp.Columns(12).Width = 40
            GrdHelp.Columns(13).Width = 40
            If gblnGSTUnit Then
                If GetPlantName() = "HILEX" Then
                    GrdHelp.Columns(16).Width = 60
                    GrdHelp.Columns(17).Width = 80
                    GrdHelp.Columns(18).Width = 50
                    GrdHelp.Columns(19).Width = 50
                    GrdHelp.Columns(20).Width = 50
                    GrdHelp.Columns(21).Width = 50
                    GrdHelp.Columns(22).Width = 50
                Else
                    GrdHelp.Columns(14).Width = 60
                    GrdHelp.Columns(15).Width = 80
                    GrdHelp.Columns(16).Width = 50
                    GrdHelp.Columns(17).Width = 50
                    GrdHelp.Columns(18).Width = 50
                    GrdHelp.Columns(19).Width = 50
                    GrdHelp.Columns(20).Width = 50
                End If
            End If
            ' GrdHelp.DefaultCellStyle.BackColor = Color.DarkKhaki
            'Dim j As Integer
            'For j = 0 To GrdHelp.ColumnCount - 1
            '    GrdHelp.Columns(j).Width = 4

            'Next
            Dim colno As Integer
            For i = 1 To dt.Columns.Count - 1
                Me.cboSearchBy.Items.Add(dt.Columns(i).ColumnName)
                GrdHelp.Columns(i).ReadOnly = True
            Next
            If mblnFtsEnabled = True Then
                For i = 0 To dt.Rows.Count - 1
                    If GrdHelp.Rows(i).Cells(14).Value = True And GrdHelp.Rows(i).Cells(15).Value = True Then
                        For colno = 1 To dt.Columns.Count - 1
                            GrdHelp.Rows(i).Cells(colno).Style.BackColor = Color.Bisque
                        Next

                    ElseIf GrdHelp.Rows(i).Cells(14).Value = True And GrdHelp.Rows(i).Cells(15).Value = False Then
                        For colno = 1 To dt.Columns.Count - 1
                            GrdHelp.Rows(i).Cells(colno).Style.BackColor = Color.LightSkyBlue
                        Next
                    End If
                Next
            End If

            cboSearchBy.SelectedIndex = 0
            GrdHelp.ClearSelection()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'SetBackGroundColorNew(Me, True)
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Me.Height = Me.CmdOK.Top + Me.CmdOK.Height + 100
        mblnFtsEnabled = FTS_FUNCTIONALITY()
        FillHelp()
        If mblnFtsEnabled = True Then
            FTS_COLOURSYMBOL.Visible = True
        Else
            FTS_COLOURSYMBOL.Visible = False
        End If
        '101188073 Start
        _itemCode = String.Empty
        '101188073 End
        System.Windows.Forms.Cursor.Current = Cursors.Default
        Exit Sub
    End Sub

    Private Sub GrdHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GrdHelp.Click
        Dim varItemCode As Object
        Dim varKanbanNo As Object
        Dim blnflag As Boolean
        Dim blnflag1 As Boolean
        Dim intSubItem As Short
        Dim StrItemCode As String
        Dim varToolCost As Object

        '''***** Addded by ashutosh on 26-12-2005, Issue id:16685
        Dim intToolCost As Short
        Dim intOtherItemToolCost As Short
        Dim intSubItem1 As Short
        Dim blnOtherRecord As Boolean
        Dim blnFTSitemstatus As Boolean
        Dim blnFTSbarcode As Boolean
        Dim colno As Integer
        mblnFtsEnabled = FTS_FUNCTIONALITY()
        'FillHelp()
        If mblnFtsEnabled = True Then
            For i = 0 To dt.DefaultView.Count - 1
                If GrdHelp.Rows(i).Cells(14).Value = True And GrdHelp.Rows(i).Cells(15).Value = True Then
                    For colno = 1 To dt.Columns.Count - 1
                        GrdHelp.Rows(i).Cells(colno).Style.BackColor = Color.Bisque
                    Next

                ElseIf GrdHelp.Rows(i).Cells(14).Value = True And GrdHelp.Rows(i).Cells(15).Value = False Then
                    For colno = 1 To dt.Columns.Count - 1
                        GrdHelp.Rows(i).Cells(colno).Style.BackColor = Color.LightSkyBlue
                    Next

                End If


            Next
        End If

        'If mblnFtsEnabled = True Then
        '    With GrdHelp
        '        If .Rows(intSubItem).Cells(GridHeader.FTS_ITEM).Value = True And .Rows(intSubItem).Cells(GridHeader.FTS_BARCODE).Value = True Then
        '            GrdHelp.DefaultCellStyle.BackColor = Color.Bisque
        '        End If
        '        If .Rows(intSubItem).Cells(GridHeader.FTS_ITEM).Value = True And .Rows(intSubItem).Cells(GridHeader.FTS_BARCODE).Value = False Then
        '            GrdHelp.DefaultCellStyle.BackColor = Color.LightSkyBlue
        '        End If
        '    End With
        'End If
    End Sub

    Private Sub GrdHelp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles GrdHelp.KeyDown
        If e.KeyCode = Keys.Enter Then
            CmdOK.PerformClick()
        End If
    End Sub

    Private Sub grdHelp_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GrdHelp.RowEnter
        mRowIndex = e.RowIndex
    End Sub

    Public WriteOnly Property SqlQry() As String
        Set(ByVal value As String)
            mStrSql = value
        End Set
    End Property

    Public ReadOnly Property SelectedRows() As DataTable
        Get
            Return dtSelectedRows
        End Get
    End Property

    Private Function CreateDataTable() As DataTable
        Dim dt As New DataTable("table1")
        Dim Col As DataColumn
        Dim intCol As Integer
        Try
            For intCol = 0 To Me.GrdHelp.Columns.Count - 1
                Col = New DataColumn()
                Col.DataType = Me.GrdHelp.Columns(intCol).ValueType
                Col.ColumnName = Me.GrdHelp.Columns(intCol).Name

                dt.Columns.Add(Col)
            Next
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Nothing
        End Try
    End Function

    Public Function ShowHelpNagare(ByVal StrQry As String) As DataTable
        Dim dt As DataTable
        Try
            Dim Help As New FrmHelpNagare
            With Help
                .SqlQry = StrQry
                .ShowDialog()
                'dt = .SelectedRows
                Help = Nothing
                If Not dt Is Nothing Then
                    Return dt
                Else
                    Return Nothing
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Nothing
        End Try
    End Function

    Public Function SelectDatafromItem_Mst(Optional ByRef pstrItem As String = "", Optional ByRef intAlreadyItem As Integer = 0) As Object
        On Error GoTo ErrHandler
        Dim strItembal As String
        Dim rsItembal As ClsResultSetDB
        Dim rsbackdate_pds As ClsResultSetDB
        Dim strBackdateqry As String
        Dim intRecordCount As Integer 'To Hold Record Count
        Dim intCount As Short
        Dim blnBackDatePDS As Boolean
        Dim startingdate As String
        startingdate = getDateForDB(GetServerDate())

        strItembal = "set dateformat 'dmy' select Distinct convert(bit,0)as CHK, KanbanNo,m.Account_code,m.item_code, m.cust_drgNo as DrawingNo, m.Description, cast(m.Quantity as varchar(20)) as Quantity ,  convert(varchar(12),m.Sch_Date,103) Sch_Date, case when Sch_Time = '23:59' then '' else Sch_Time end as sch_time,m.Cust_Ref, m.amendment_no,"
        strItembal = strItembal & " cast(m.Tool_cost as varchar(20)) as Tool_cost,m.UNLOC, m.USLOC"
        If GetPlantName() = "HILEX" Then
            strItembal = strItembal & ",M.FTS_ITEM as FTS_ITEM ,M.FTS_BARCODE_TRACKING as FTS_BARCODE "
            '101188073 Start
            If gblnGSTUnit Then
                strItembal = strItembal & " ,m.ISHSNORSAC HSN_SAC_TYPE,m.HSNSACCODE HSN_SAC_CODE,m.CGSTTXRT_TYPE [CGST],m.SGSTTXRT_TYPE [SGST],m.IGSTTXRT_TYPE [IGST],m.UTGSTTXRT_TYPE [UTGST],m.COMPENSATION_CESS COMP_CESS,m.batch_code "
            End If
            '101188073 End
            strItembal = strItembal & " from vw_Enagaredtl_Help_HILEX m"
        Else
            '101188073 Start
            If gblnGSTUnit Then
                strItembal = strItembal & " ,m.ISHSNORSAC HSN_SAC_TYPE,m.HSNSACCODE HSN_SAC_CODE,m.CGSTTXRT_TYPE [CGST],m.SGSTTXRT_TYPE [SGST],m.IGSTTXRT_TYPE [IGST],m.UTGSTTXRT_TYPE [UTGST],m.COMPENSATION_CESS COMP_CESS,m.batch_code "
            End If
            '101188073 End
            strItembal = strItembal & " from vw_Enagaredtl_Help m"
        End If

        'strItembal = strItembal & " where m.UNIT_CODE='" & gstrUNITID & "' and( M.SCH_DATE >= CONVERT(CHAR(13),'" & startingdate & "' , 106)  ) and m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.location_code = b.location_code and a.UNIT_CODE = b.UNIT_CODE and a.doc_no=b.doc_no where m.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND m.kanbanNo = b.srvdino and a.cancel_flag <> 1) + (select IsNull(sum(sales_quantity),0) as sales_quantity  from printedsrv_dtl p where p.UNIT_CODE=m.UNIT_CODE AND p.UNIT_CODE='" & gstrUNITID & "' AND p.KanBan_No=m.KanBanNo)+(Select isnull(Sum(quantity),0) as sales_quantity From mkt_57F4challankanban_dtl B inner join mkt_57F4challan_hdr A on B.doc_type=A.doc_type and B.doc_no = A.doc_no AND B.UNIT_CODE = A.UNIT_CODE where B.UNIT_CODE=m.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' AND A.cancel_flag = 0 and B.Kanban_no=m.KanBanNo))"
        'strItembal = strItembal & " where ( M.SCH_DATE >= CONVERT(CHAR(13),'" & startingdate & "' , 106)  ) and m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.location_code = b.location_code and a.doc_no=b.doc_no where m.kanbanNo = b.srvdino and a.bill_flag=1 and a.cancel_flag =0) )"
        strItembal = strItembal & " where m.UNIT_CODE='" & gstrUnitId & "' and( M.SCH_DATE >= CONVERT(CHAR(13),'" & startingdate & "' , 106)  ) and m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.location_code = b.location_code and a.UNIT_CODE = b.UNIT_CODE and a.doc_no=b.doc_no where m.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE='" & gstrUnitId & "' and m.kanbanNo = b.srvdino and a.cancel_flag <>1))"
        strItembal = strItembal & " order by SCH_DATE DESC , SCH_TIME DESC , kanbanNo "
        ShowHelpNagare(strItembal)

        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

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
        Dim strkanbanno As String
        Dim strotherkanbanno As String
        Dim strItemCreation() As String
        Dim IntMaxvalue As Integer
        Dim IntLoopCounter As Integer
        Dim intinnerloop As Integer
        Dim strftsitem As String
        Dim strftsbarcode As String
        Dim strotherftsitem As String
        Dim strotherftsbarcode As String

        Me.LBLITEMCREATION.Text = ""
        With GrdHelp
            For intSubItem = 0 To .Rows.Count - 1


                If .Rows(intSubItem).Cells(GridHeader.Mark).Value = True Then
                    If Not blnOtherRecord Then
                        '.Col = GridHeader.AccountCode
                        'strAccountCode = .Rows(intSubItem).Cells(GridHeader.AccountCode).Value
                        strAccountCode = .Rows(intSubItem).Cells("account_code").Value
                        '.Col = GridHeader.CustRef
                        strCustRef = .Rows(intSubItem).Cells("cust_ref").Value
                        ' .Col = GridHeader.AmendmentNo
                        StrAmendmentNo = .Rows(intSubItem).Cells("amendment_no").Value
                        '.Col = GridHeader.ItemCode
                        StrItemCode = .Rows(intSubItem).Cells("item_code").Value
                        Me.LBLITEMCREATION.Text = StrItemCode

                        strkanbanno = .Rows(intSubItem).Cells("kanbanno").Value
                        strSalesTaxType = Trim(Find_Value("select SalesTax_Type from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and amendment_no='" & StrAmendmentNo & "'"))
                        strSOType = Trim(Find_Value("select PO_Type from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and amendment_no='" & StrAmendmentNo & "'"))
                        strExciseType = Trim(Find_Value("select Excise_duty from cust_ord_dtl where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and Amendment_no ='" & StrAmendmentNo & "' and item_code ='" & StrItemCode & "' and cust_drgNo='" & StrItemCode & "'"))

                        If mblnFtsEnabled = True Then
                            strftsitem = .Rows(intSubItem).Cells("FTS_ITEM").Value
                            strftsbarcode = .Rows(intSubItem).Cells("FTS_Barcode").Value
                        End If

                        blnOtherRecord = True
                    Else
                        '.Col = GridHeader.AccountCode
                        strOtherAccountCode = .Rows(intSubItem).Cells("Account_Code").Value
                        ' .Col = GridHeader.CustRef
                        strOtherCustRef = .Rows(intSubItem).Cells("Cust_Ref").Value
                        '.Col = GridHeader.AmendmentNo
                        strOtherAmendmentNo = .Rows(intSubItem).Cells("Amendment_No").Value
                        '.Col = GridHeader.ItemCode
                        strOtherItemCode = .Rows(intSubItem).Cells("Item_Code").Value

                        'strItemCreation(intSubItem) = StrItemCode.Trim.ToString
                        Me.LBLITEMCREATION.Text = Me.LBLITEMCREATION.Text + "," + strOtherItemCode
                        strotherkanbanno = .Rows(intSubItem).Cells("KanbanNo").Value
                        strOtherSalesTaxType = Trim(Find_Value("select SalesTax_Type from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and amendment_no='" & strOtherAmendmentNo & "'"))
                        strOtherSOType = Trim(Find_Value("select PO_Type from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and amendment_no='" & strOtherAmendmentNo & "'"))
                        strOtherExciseType = Trim(Find_Value("select Excise_duty from cust_ord_dtl where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and Amendment_no ='" & strOtherAmendmentNo & "' and item_code ='" & strOtherItemCode & "' and cust_drgNo='" & strOtherItemCode & "'"))
                        If mblnFtsEnabled = True Then
                            strotherftsitem = .Rows(intSubItem).Cells("FTS_ITEM").Value
                            strotherftsbarcode = .Rows(intSubItem).Cells("FTS_Barcode").Value
                        End If

                        If UCase(strAccountCode) <> UCase(strOtherAccountCode) Then
                            strMessage = "Two or more SOs of different Customers are not allowed." & vbCrLf
                            strMessage = strMessage & "1. " & strAccountCode & " -> " & strCustRef & vbCrLf
                            strMessage = strMessage & "2. " & strOtherAccountCode & " -> " & strOtherCustRef
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If

                        If UCase(strSOType) <> UCase(strOtherSOType) And gblnGSTUnit = False Then
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
                            'If UCase(StrItemCode) = UCase(strOtherItemCode) Then
                            '    strMessage = "Two or more items of different Kanbanno are not allowed." & vbCrLf
                            '    strMessage = strMessage & "1. " & StrItemCode & " -> " & strkanbanno & vbCrLf
                            '    strMessage = strMessage & "2. " & strOtherItemCode & " -> " & strotherkanbanno
                            '    MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            '    ValidateData = False
                            '    Exit Function
                            'End If

                            If UCase(strExciseType) <> UCase(strOtherExciseType) Then
                                strMessage = "Two or more SOs can not have different Excise Rate." & vbCrLf
                                strMessage = strMessage & "1. " & strCustRef & " -> " & strExciseType & vbCrLf
                                strMessage = strMessage & "2. " & strOtherCustRef & " -> " & strOtherExciseType
                                MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                                ValidateData = False
                                Exit Function
                            End If
                        End If '101188073
                        If frmMKTTRN0035.OptNormalDispatch.Checked = True Then
                            'If mblnFtsSpareDispatch = False Then
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
            strItemCreation = Me.LBLITEMCREATION.Text.Split(",")
            IntMaxvalue = UBound(strItemCreation)
            If IntMaxvalue > 0 Then
                For IntLoopCounter = 0 To UBound(strItemCreation)

                    For intinnerloop = IntLoopCounter + 1 To UBound(strItemCreation)
                        If strItemCreation(IntLoopCounter).Trim.ToUpper = strItemCreation(intinnerloop).Trim.ToUpper Then
                            strMessage = "Two or more Same items of different Kanbanno are not allowed." & vbCrLf
                            strMessage = strMessage & "1. " & strItemCreation(IntLoopCounter).Trim.ToUpper & vbCrLf
                            strMessage = strMessage & "2. " & strItemCreation(intinnerloop).Trim.ToUpper
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If
                    Next
                Next
            End If
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

    Private Sub CmdOK_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdOK.Click
        Dim Row As DataRow
        Dim intCell As Integer
        mstrItemText = ""
        Dim intSubItem As Short
        Dim strKanabanNo As String
        Dim StrItemCode As String
        Dim strDrawingNo As String
        Dim strDescription As String
        Dim dblQuantity As Double
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim strSchDate As String
        Dim strSchTime As String
        Dim strunLoc As String
        Dim strUSLoc As String
        Dim strAccountCode As String
        Dim sqlCmd As SqlCommand
        Dim dsTax As DataSet
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
        Dim strBatchno As String
        '101188073 End
        Try

            If Not ValidateData() Then Exit Sub
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
            '''''''''''''''''''Added by geetanjali to support Multi unit for HILEX''''''''''''''''
            If Me.GrdHelp.RowCount <= 0 Then Exit Sub
            dtSelectedRows = New DataTable
            dtSelectedRows = CreateDataTable()
            Row = dtSelectedRows.NewRow
            For intCell = 0 To Me.GrdHelp.CurrentRow.Cells.Count - 1
                Row(intCell) = Me.GrdHelp.CurrentRow.Cells(intCell).Value
            Next
            dtSelectedRows.Rows.Add(Row)
            With GrdHelp
                For intSubItem = 0 To .Rows.Count - 1
                    flag = False
                    If .Rows(intSubItem).Cells(GridHeader.Mark).Value = True Then
                        '101188073 Start
                        If gblnGSTUnit Then
                            If strItem IsNot Nothing AndAlso strItem.Count > 0 Then
                                For arrayIndex As Integer = 0 To strItem.Count - 1
                                    If Trim(.Rows(intSubItem).Cells(GridHeader.ItemCode).Value) = Trim(strItem(arrayIndex).ToString()) Then
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
                        strKanabanNo = .Rows(intSubItem).Cells("KanbanNo").Value
                        StrItemCode = .Rows(intSubItem).Cells("Item_Code").Value
                        dblQuantity = .Rows(intSubItem).Cells("Quantity").Value
                        strCustRef = .Rows(intSubItem).Cells("Cust_Ref").Value
                        StrAmendmentNo = .Rows(intSubItem).Cells("Amendment_No").Value
                        strSchDate = .Rows(intSubItem).Cells("Sch_Date").Value
                        strSchTime = .Rows(intSubItem).Cells("SCh_Time").Value
                        strunLoc = .Rows(intSubItem).Cells("UNLoc").Value
                        strUSLoc = .Rows(intSubItem).Cells("USLOC").Value
                        strAccountCode = .Rows(intSubItem).Cells("Account_Code").Value
                        strDrawingNo = .Rows(intSubItem).Cells("DrawingNo").Value
                        strBatchno = .Rows(intSubItem).Cells("batch_Code").Value
                        Dim blnftsitem As Boolean
                        Dim blnftsBarcode As Boolean

                        blnftsitem = CBool(Trim(Find_Value("select FTS_ITEM from item_mst where UNIT_CODE = '" & gstrUnitId & "' AND item_code='" & StrItemCode & "'")))
                        blnftsBarcode = CBool(Trim(Find_Value("select FTS_BARCODE_TRACKING from item_mst where UNIT_CODE = '" & gstrUnitId & "' AND item_code='" & StrItemCode & "'")))
                        frmMKTTRN0035.FTSItem = blnftsitem
                        frmMKTTRN0035.FTSBarcode = blnftsBarcode
                        '101188073 Start
                        If gblnGSTUnit Then
                            hsnsacType = .Rows(intSubItem).Cells("HSN_SAC_TYPE").Value
                            hsnsacCode = .Rows(intSubItem).Cells("HSN_SAC_CODE").Value
                            cgstType = .Rows(intSubItem).Cells("CGST").Value
                            sgstType = .Rows(intSubItem).Cells("SGST").Value
                            igstType = .Rows(intSubItem).Cells("IGST").Value
                            utgstType = .Rows(intSubItem).Cells("UTGST").Value
                            CompensationCessType = .Rows(intSubItem).Cells("COMP_CESS").Value
                            sqlCmd = New SqlCommand("SELECT CGST_PERCENT,SGST_PERCENT,IGST_PERCENT,UTGST_PERCENT,COMPENSATION_CESS_PERCENT FROM dbo.UDF_GST_TAX_RATE_PERCENT('" & gstrUnitId & "','" & .Rows(intSubItem).Cells("CGST").Value & "','" & .Rows(intSubItem).Cells("SGST").Value & "','" & .Rows(intSubItem).Cells("IGST").Value & "','" & .Rows(intSubItem).Cells("UTGST").Value & "','" & .Rows(intSubItem).Cells("COMP_CESS").Value & "')")
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
                            mstrItemText = mstrItemText & strKanabanNo & "|" & StrItemCode & "|" & strDrawingNo & "|" & strDescription & "|" & dblQuantity & "|" & strCustRef & "|" & StrAmendmentNo & "|" & strSchDate & "|" & strSchTime & "|" & strunLoc & "|" & strUSLoc & "|" & strAccountCode & "|" & hsnsacType & "|" & hsnsacCode & "|" & cgstType & "|" & cgstPercent & "|" & sgstType & "|" & sgstPercent & "|" & igstType & "|" & igstPercent & "|" & utgstType & "|" & utgstPercent & "|" & CompensationCessType & "|" & CompensationCessPercent & "|" & strBatchno & "^"
                        Else
                            mstrItemText = mstrItemText & strKanabanNo & "|" & StrItemCode & "|" & strDrawingNo & "|" & strDescription & "|" & dblQuantity & "|" & strCustRef & "|" & StrAmendmentNo & "|" & strSchDate & "|" & strSchTime & "|" & strunLoc & "|" & strUSLoc & "|" & strAccountCode & "|" & strBatchno & "^"
                        End If
                        '101188073 End
                    End If
                Next intSubItem
            End With
            If Len(mstrItemText) = 0 Then
                Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Me.GrdHelp.Focus()
                Exit Sub
            End If
            Me.Dispose()

            Exit Sub

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdCancel.Click
        Me.Close()
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
            With GrdHelp
                dtGST.Columns.Add("ITEM_CODE", GetType(String))
                dtGST.Columns.Add("UNIT_CODE", GetType(String))
                dtGSTHSN.Columns.Add("ITEM_CODE", GetType(String))
                dtGSTHSN.Columns.Add("UNIT_CODE", GetType(String))
                For rowIndex As Integer = 0 To .Rows.Count - 1
                    If .Rows(rowIndex).Cells(GridHeader.Mark).Value Then
                        drGST = dtGST.NewRow()
                        drGSTHSN = dtGSTHSN.NewRow()
                        drGST("ITEM_CODE") = .Rows(rowIndex).Cells("item_code").Value
                        drGST("UNIT_CODE") = gstrUnitId
                        drGSTHSN("ITEM_CODE") = .Rows(rowIndex).Cells("item_code").Value
                        drGSTHSN("UNIT_CODE") = gstrUnitId
                        dtGST.Rows.Add(drGST)
                        If Len(Trim(.Rows(rowIndex).Cells("HSN_SAC_CODE").Value)) = 0 Then
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

    Private Sub GrdHelp_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GrdHelp.CellContentClick

    End Sub
End Class
