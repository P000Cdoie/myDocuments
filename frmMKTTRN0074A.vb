Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports System.Data
Friend Class frmMKTTRN0074A
    Inherits System.Windows.Forms.Form
#Region "Comments"
    '***************************************************************************************
    'Copyright       : MIND Ltd.
    'Module          : frmMKTTRN0074A - Intra Material Transfer Note Help Form
    'Author          : Parveen Kumar
    'Creation Date   : 04 Jul 2012
    'Description     : Intra Material Transfer Note Help Form
    '***************************************************************************************
    'Revised By       - Prashant rajpal
    'Revision Date    - 03 feb 2014
    'Issue Id         - 10488279
    'Description      - RAW Material and BOP is incorporated in this form
    '---------------------------------------------------------------------------

#End Region
#Region "Form level variable Declarations"
    Dim mCtlHdrItemCode As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrDrawingNo As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrDescription As System.Windows.Forms.ColumnHeader
    Dim intCheckCounter As Short = 0
    Dim mListItemUserId As System.Windows.Forms.ListViewItem
    Dim blnExpinv As Boolean = False
    Dim intIteminSp As Short = 0
    Dim mStrSql As String = String.Empty
#End Region
#Region "Form Control Events"
    Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOk.Click
        Try
            mstrItemText = "" : intCheckCounter = intIteminSp
            Dim intSubItem As Short
            Dim gobjDB As ClsResultSetDB
            Dim blnMoreThan7ItemInInvoice As Boolean
            gobjDB = New ClsResultSetDB
            blnMoreThan7ItemInInvoice = False

            gobjDB.GetResult("Select MoreThan7ItemInInvoice from sales_parameter where unit_code='" & gstrUNITID & "'")
            If gobjDB.GetValue("MoreThan7ItemInInvoice") = True Then
                blnMoreThan7ItemInInvoice = True
            Else
                blnMoreThan7ItemInInvoice = False
            End If
            '--------------------------------------------------------------------------
            For intSubItem = 0 To lvwItemCode.Items.Count - 1
                If Me.lvwItemCode.Items.Item(intSubItem).Checked = True Then
                    intCheckCounter = intCheckCounter + 1
                    If blnExpinv = False Then
                        If blnMoreThan7ItemInInvoice = True Then
                            gobjDB = New ClsResultSetDB
                            gobjDB.GetResult("select MaximumItemsInInvoices from sales_parameter where unit_code='" & gstrUNITID & "'")
                            If intCheckCounter > gobjDB.GetValue("MaximumItemsInInvoices") Then
                                MsgBox("No. Of Items Selected Should not be greater than " & gobjDB.GetValue("MaximumItemsInInvoices") & "", MsgBoxStyle.Information, "empower")
                                mstrItemText = ""
                                Exit Sub
                            End If
                        End If
                    Else
                        gobjDB = New ClsResultSetDB
                        gobjDB.GetResult("Select EOU_Flag,company_code from Company_Mst where unit_code='" & gstrUNITID & "'")
                        If gobjDB.GetValue("EOU_Flag") = False Then
                            gobjDB.ResultSetClose()
                            gobjDB = New ClsResultSetDB
                            gobjDB.GetResult("Select MoreThan7ItemInInvoice from sales_parameter where unit_code='" & gstrUNITID & "'")
                            If gobjDB.GetValue("MoreThan7ItemInInvoice") = False Then
                                gobjDB.ResultSetClose()
                                gobjDB = Nothing
                                If intCheckCounter > 7 Then
                                    MsgBox("No. Of Items Selected Should be Less than 7", MsgBoxStyle.Information, "empower")
                                    mstrItemText = ""
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    mstrItemText = mstrItemText & "'" & Trim(Me.lvwItemCode.Items.Item(intSubItem).SubItems(1).Text) & "',"
                End If
            Next intSubItem
            If Len(mstrItemText) = 0 Then
                Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Me.lvwItemCode.Focus()
                Exit Sub
            End If
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub lvwItemCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvwItemCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    CmdOk.Focus()
            End Select
            GoTo EventExitSub
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub optDescription_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDescription.CheckedChanged
        If eventSender.Checked Then
            Try
                With lvwItemCode
                    .Sort()
                    ListViewColumnSorter.SortListView(lvwItemCode, 2, SortOrder.Ascending)
                    .Sorting = System.Windows.Forms.SortOrder.Ascending
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Private Sub OptItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItemCode.CheckedChanged
        If eventSender.Checked Then
            Try
                With lvwItemCode
                    .Sort()
                    ListViewColumnSorter.SortListView(lvwItemCode, 0, SortOrder.Ascending)
                    .Sorting = System.Windows.Forms.SortOrder.Ascending
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Private Sub optPartNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPartNo.CheckedChanged
        If eventSender.Checked Then
            Try
                With lvwItemCode
                    .Sort()
                    ListViewColumnSorter.SortListView(lvwItemCode, 1, SortOrder.Ascending)
                    .Sorting = System.Windows.Forms.SortOrder.Ascending
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
    End Sub
#End Region
#Region "Sub Routines"
    Private Sub SearchItem()
        Try
            Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
            If optDescription.Checked = True Then
                itmFound = SearchText((txtsearch.Text), optDescription, lvwItemCode, "2")
            Else
                itmFound = SearchText((txtsearch.Text), optPartNo, lvwItemCode)
            End If
            If itmFound Is Nothing Then ' If no match,
                Exit Sub
            Else
                itmFound.EnsureVisible() ' Scroll ListView to show found ListItem.
                itmFound.Selected = True ' Select the ListItem.
                lvwItemCode.Enabled = True
                If Len(txtsearch.Text) > 0 Then itmFound.Font = VB6.FontChangeBold(itmFound.Font, True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub AddColumnsInListView()
        Try
            With Me.lvwItemCode
                mCtlHdrItemCode = .Columns.Add("")
                mCtlHdrItemCode.Text = "Item Code"
                mCtlHdrDrawingNo = .Columns.Add("")
                mCtlHdrDrawingNo.Text = "Drawing No."
                mCtlHdrDescription = .Columns.Add("")
                mCtlHdrDescription.Text = "Description"
                mCtlHdrDescription = .Columns.Add("")
                mCtlHdrDescription.Text = "Tariff Code"
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
#End Region
#Region "FormEvents"
    
#End Region
    

    '###########################################################################################################
    '##########################ONLY FOR INTRA MATERIAL TRANSFER NOTE ##########################################
    Public Function CheckSoItems(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, Optional ByRef pstrItemNotin As String = "") As String
        Try
            Dim SqlCmd As SqlCommand = Nothing
            Dim SqlDR As SqlDataReader = Nothing
            Dim SqlAdp As SqlDataAdapter = Nothing
            Dim sqlDataSet As DataSet = Nothing

            mStrSql = " delete from TEMP_INTRAMAT where ip_address = '" & gstrIpaddressWinSck.ToString.Trim & "' and unit_code = '" & gstrUNITID.ToString.Trim & "'"
            SqlCmd = New SqlCommand(mStrSql, SqlConnectionclass.GetConnection)
            SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
            SqlCmd = Nothing
            'Issue ID-10488279
            mStrSql = "  INSERT INTO TEMP_INTRAMAT(ACCOUNT_CODE,ITEM_CODE,ITEM_GLGRP,ITEM_GLCODE,ITEM_SLCODE, " & _
                    " Active_Flag,STATUS,ITEM_MAIN_GRP, " & _
                    " LOCATION_CODE,CUR_BAL,IP_Address,Unit_Code) " & _
                    " SELECT A.Account_Code, C.ITEM_CODE,D.GlGrp_code,INV.invGld_glCode, " & _
                    " ISNULL(INV.invGld_slCode,'') AS invGld_slCode,c.Active_Flag,D.Status,D.Item_Main_Grp , " & _
                    " CASE WHEN D.Item_Main_Grp IN ('F','S') THEN '01P1' " & _
                    " WHEN  D.Item_Main_Grp IN('R','C') THEN '01M1' END AS LOCATION_CODE,0,'" & gstrIpaddressWinSck & "','" & gstrUNITID & "' " & _
                  " FROM Cust_Ord_hdr a INNER JOIN Cust_ord_dtl c " & _
                    " ON A.Account_Code = C.Account_Code AND A.Cust_Ref = C.Cust_Ref  " & _
                    " AND A.Amendment_No = C.Amendment_No AND A.UNIT_CODE = C.UNIT_CODE " & _
                    " INNER JOIN Item_Mst d ON C.Item_Code = D.Item_Code AND A.UNIT_CODE = D.UNIT_CODE " & _
                    " INNER JOIN fin_InvGLGrpDTL INV ON D.UNIT_CODE = INV.UNIT_CODE AND D.GlGrp_code = INV.invGld_invGLGrpId " & _
                    " WHERE A.UNIT_CODE = '" & gstrUNITID.ToString.Trim & "'  " & _
                    " AND A.ACCOUNT_CODE = '" & pstrCustno.ToString.Trim & "' " & _
                    " AND c.Active_Flag = 'A' AND D.Status='A' And A.Cust_ref='" & pstrRefNo & "' AND A.Amendment_No ='" & pstrAmmNo & "' " & _
                    " AND INV.invGld_prpsCode = 'StockTrans'"
            'Issue ID-10488279
            If pstrItemNotin.ToString.Trim.Length > 0 Then
                mStrSql = mStrSql & " and c.Item_code not in ( " & pstrItemNotin & ")"
            End If
            SqlCmd = New SqlCommand(mStrSql, SqlConnectionclass.GetConnection)
            SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
            SqlCmd = Nothing

            mStrSql = " UPDATE T SET T.CUR_BAL = TBL1.CUR_BAL " & _
                " FROM TEMP_INTRAMAT T INNER JOIN " & _
                " (SELECT ITEM_CODE,LOCATION_CODE,CUR_BAL,UNIT_CODE FROM Itembal_Mst WHERE UNIT_CODE='" & gstrUNITID.ToString.Trim & "')TBL1 " & _
                " ON T.ITEM_CODE = TBL1.item_code AND T.LOCATION_CODE = TBL1.location_code " & _
                " WHERE TBL1.UNIT_CODE = '" & gstrUNITID.ToString.Trim & "' and t.IP_Address = '" & gstrIpaddressWinSck.ToString.Trim & "'"

            SqlCmd = New SqlCommand(mStrSql, SqlConnectionclass.GetConnection)
            SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
            SqlCmd = Nothing

            Dim strDate As String = String.Empty
            Dim Validyrmon As String = String.Empty
            Dim effectyrmon As String = String.Empty
            Dim validMon As String = String.Empty
            Dim effectMon As String = String.Empty

            strDate = VB6.Format(GetServerDate(), gstrDateFormat)
            Me.lvwItemCode.Items.Clear() 'initially clear all items in the listview
            mStrSql = "Select effectMon=convert(char(2),month(effect_date)),effectYr=convert(char(4),Year(effect_date)), " & _
            " validMon=convert(char(2),month(Valid_date)),validYr=convert(char(4),year(Valid_date)) " & _
            " from Cust_Ord_hdr where unit_code='" & gstrUNITID & "' and " & _
            " Account_Code='" & pstrCustno.ToString.Trim & "' and Cust_Ref='" & pstrRefNo.ToString.Trim & "' " & _
            " and Amendment_No='" & pstrAmmNo.ToString.Trim & "' and Active_Flag = 'A' "
            SqlCmd = New SqlCommand
            With SqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = mStrSql
                SqlDR = SqlCmd.ExecuteReader()
                If SqlDR.Read = True Then
                    validMon = CStr(Month(GetServerDate))
                    If CDbl(validMon) < 10 Then
                        validMon = "0" & validMon
                    End If
                    Validyrmon = Year(GetServerDate) & validMon
                    effectMon = SqlDR("EffectMon")
                    If CDbl(effectMon) < 10 Then
                        effectMon = "0" & effectMon
                    End If
                    effectyrmon = SqlDR("effectYr") & effectMon
                Else
                    If SqlDR.IsClosed = False Then SqlDR.Close()
                    SqlCmd.Connection.Close()
                    SqlCmd.Dispose()

                    mStrSql = "Select effect_date,Valid_date from Cust_Ord_hdr " & _
                            " where unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(pstrCustno) & "'  " & _
                            " and Cust_Ref='" & Trim(pstrRefNo) & "' and Amendment_No='" & Trim(pstrAmmNo) & "'  " & _
                            " and Active_flag ='A'"
                    SqlCmd = New SqlCommand
                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandType = CommandType.Text
                        .CommandText = mStrSql
                        SqlDR = SqlCmd.ExecuteReader()
                        If SqlDR.Read = True Then
                            Validyrmon = SqlDR("valid_date")
                            effectyrmon = SqlDR("Effect_date")
                        Else
                            MessageBox.Show("No Items for intra material transfer in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items at related location." & vbCrLf & "3. Check Marketing Schedule for intra material transfer Goods in SO.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            Return String.Empty
                            Exit Function
                        End If
                        If SqlDR.IsClosed = False Then SqlDR.Close()
                        SqlCmd.Connection.Close()
                        SqlCmd.Dispose()
                    End With

                End If
                If SqlDR.IsClosed = False Then SqlDR.Close()
                SqlCmd.Connection.Close()
                SqlCmd.Dispose()

                Dim intRowCount As Integer = 0
                mStrSql = " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code " & _
                " from Cust_Ord_hdr a,MonthlyMktSchedule b, " & _
                " Cust_ord_dtl c,Item_Mst d  " & _
                "where a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code  " & _
                " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code  " & _
                " And c.Active_Flag ='A' and a.account_code = b.Account_code and c.Cust_drgNo = b.Cust_drgNo  " & _
                " and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "'  " & _
                " and a.unit_code='" & gstrUNITID & "' and a.Cust_Ref='" & Trim(pstrRefNo) & "'  " & _
                " and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1 and b.Schedule_flag =1  " & _
                " and b.Year_Month =  " & Validyrmon & " " & _
                " and b.Item_Code in(Select a.Item_code from Item_MSt a inner join TEMP_INTRAMAT b " & _
                " on a.unit_code = b.unit_code and a.Item_code = b.Item_code " & _
                " where b.ip_address = '" & gstrIpaddressWinSck & "' and  a.unit_code='" & gstrUNITID & "' " & _
                " and b.Cur_bal >0  and a.hold_flag =0 and a.Status = 'A' "
                If pstrItemNotin.ToString.Trim.Length > 0 Then
                    mStrSql = mStrSql & " and a.Item_code not in(" & pstrItemNotin.ToString.Trim & ")) "
                Else
                    mStrSql = mStrSql & ")"
                End If
                mStrSql = mStrSql & " UNION  " & _
                " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code  " & _
                " from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d   " & _
                " where a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code  " & _
                " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code " & _
                " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code  " & _
                " and b.status = 1 And c.Active_Flag ='A' and a.Account_Code='" & Trim(pstrCustno) & "'  " & _
                " and a.unit_code='" & gstrUNITID & "' and a.Cust_Ref='" & Trim(pstrRefNo) & "'  " & _
                " and a.Amendment_No='" & Trim(pstrAmmNo) & "'  " & _
                " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "'  " & _
                " And  b.Trans_Date <= '" & getDateForDB(strDate) & "'   " & _
                " And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "' " & _
                " and b.Item_Code in(Select a.Item_code from Item_MSt a inner join TEMP_INTRAMAT b " & _
                " on a.unit_code = b.unit_code and a.Item_code = b.Item_code " & _
                " where b.ip_address = '" & gstrIpaddressWinSck & "' and  a.unit_code='" & gstrUNITID & "' " & _
                " and b.Cur_bal >0  and a.hold_flag =0 and a.Status = 'A' "
                If pstrItemNotin.ToString.Trim.Length > 0 Then
                    mStrSql = mStrSql & " and a.Item_code not in( " & pstrItemNotin.ToString.Trim & "))"
                Else
                    mStrSql = mStrSql & ")"
                End If

                SqlAdp = New SqlDataAdapter(mStrSql, SqlConnectionclass.GetConnection)
                sqlDataSet = New DataSet
                SqlAdp.Fill(sqlDataSet)

                If sqlDataSet.Tables(0).Rows.Count > 0 Then
                    For intRowCount = 0 To sqlDataSet.Tables(0).Rows.Count - 1
                        mListItemUserId = Me.lvwItemCode.Items.Add(sqlDataSet.Tables(0).Rows(intRowCount)("Item_code").ToString.Trim)
                        If mListItemUserId.SubItems.Count > 1 Then
                            mListItemUserId.SubItems(1).Text = sqlDataSet.Tables(0).Rows(intRowCount)("Cust_Drgno").ToString.Trim
                        Else
                            mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, sqlDataSet.Tables(0).Rows(intRowCount)("Cust_Drgno").ToString.Trim))
                        End If
                        If mListItemUserId.SubItems.Count > 2 Then
                            mListItemUserId.SubItems(2).Text = sqlDataSet.Tables(0).Rows(intRowCount)("Cust_Drg_Desc").ToString.Trim
                        Else
                            mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, sqlDataSet.Tables(0).Rows(intRowCount)("Cust_Drg_Desc").ToString.Trim))
                        End If
                        If mListItemUserId.SubItems.Count > 3 Then
                            mListItemUserId.SubItems(3).Text = sqlDataSet.Tables(0).Rows(intRowCount)("Tariff_Code").ToString.Trim
                        Else
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, sqlDataSet.Tables(0).Rows(intRowCount)("Tariff_Code").ToString.Trim))
                        End If
                    Next
                Else
                    MessageBox.Show("No Items for intra material transfer in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items at related location." & vbCrLf & "3. Check Marketing Schedule for intra material transfer Goods in SO.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return String.Empty
                    Exit Function
                End If
                SqlCmd.Dispose()
                SqlCmd = Nothing
                Me.ShowDialog()
                Return mstrItemText
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try

    End Function

    Private Sub frmMKTTRN0074A_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
    '###########################################################################################################
    '################################# END HERE ################################################################

    Private Sub frmMKTTRN0074A_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            SetBackGroundColorNew(Me, True)
            Call AddColumnsInListView()

            'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
            'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
            optPartNo.Checked = True
            lvwItemCode.FullRowSelect = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class