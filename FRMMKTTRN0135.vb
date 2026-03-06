'--------------------------------------------------------------------------------------------------
'COPYRIGHT      :   MIND
'CREATED BY     :   VINOD SINGH
'CREATED DATE   :   19/02/2018
'SCREEN         :   DISPATCH ADVICE AGAINST PENDING SCHEDULE
'PURPOSE        :   GENERATES PICK LIST AGAINST ITEM-CUSTOMER-PENDING SCHEDULE
'ISSUE ID       :   101462526 
'--------------------------------------------------------------------------------------------------
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Collections.Generic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared


Public Class FRMMKTTRN0135
    Dim intSave As Integer = 0
    Dim mintFormIndex As Integer
    Dim dtAuto As DataTable
    Dim blnCheckPackQtymultiple_onDA As Boolean = False

    Private Enum enmColItem
        Status = 0
        SO
        AmendmentNo
        ItemCode
        OPENSO

    End Enum

#Region "Form level Events"
    Private Sub form_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub form_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub form_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub form_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim keyascii As Short = Asc(e.KeyChar)
        Try
            If keyascii = 39 Then
                keyascii = 0
            End If
            e.KeyChar = Chr(keyascii)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain1, ctlHeader, grpButton, 500)
            Me.MdiParent = mdifrmMain
            'mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            SetItemGridsHeader()
            'btnGrp1.ShowButtons(True, False, True, False)
            'EnableControls(False, Me, True)           
            RefreshScreen()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Methods"


    Private Sub RefreshScreen()
        Try
            txtCustItem.Text = ""
            lblCusttemName.Text = ""
            txtcustomerhelp.Text = ""
            lblCustName.Text = ""
            txtItem.Text = ""
            lblItemName.Text = ""
            SetItemGridsHeader()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


#End Region



    Private Sub cmdShowSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowSaleOrder.Click
        ShowSalesOrderDATA()
    End Sub

    Private Sub ShowSalesOrderDATA()
        Try
            Dim StrSql As String = ""
            Dim i As Integer = 0
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf txtItem.Text = "" Then
                MsgBox("Select Item Code  First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf txtCustItem.Text = "" Then
                MsgBox("Select Customer Drg No. First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            Using sqlCmd1 As SqlCommand = New SqlCommand
                With sqlCmd1
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_GET_SALEORDER"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 8).Value = txtcustomerhelp.Text
                    .Parameters.Add("@NEWITEM", SqlDbType.VarChar, 16).Value = txtItem.Text
                    .Parameters.Add("@CUST_DRGNO ", SqlDbType.VarChar, 50).Value = txtCustItem.Text
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 25).Value = gstrIpaddressWinSck
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                    If Convert.ToString(.Parameters("@MESSAGE").Value).Length > 0 Then
                        MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), "eMPro", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    End If
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd1)
                    SetItemGridsHeader()
                    StrSql = "select DISTINCT CUST_REF,AMENDMENT_NO,case when OPENSO=1 then 'Open' else 'Close' end as OPENSO  from TMP_INTERNAL_ITEMSO_LINKAGE WHERE UNIT_CODE='" & gstrUNITID & "' and CUSTOMER_CODE='" & txtcustomerhelp.Text & "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                    Dim dt As DataTable = SqlConnectionclass.GetDataTable(StrSql)
                    If dt.Rows.Count > 0 Then
                        dgvSODetail.Rows.Clear()
                        dgvSODetail.Rows.Add(dt.Rows.Count)
                        For Each dr As DataRow In dt.Rows
                            dgvSODetail.Rows(i).Cells(enmColItem.Status).Value = False
                            dgvSODetail.Rows(i).Cells(enmColItem.SO).Value = dr("CUST_REF")
                            dgvSODetail.Rows(i).Cells(enmColItem.AmendmentNo).Value = dr("AMENDMENT_NO")
                            dgvSODetail.Rows(i).Cells(enmColItem.OPENSO).Value = dr("OPENSO")
                            'dgvSODetail.Rows(i).Cells(enmColItem.ItemCode).Value = dr("ITEM_CODE")
                            i += 1
                        Next
                    Else
                        MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    End If

                End With
            End Using
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Sub SetItemGridsHeader()
        Try

            dgvSODetail.Columns.Clear()

            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = "Select"

            dgvSODetail.Columns.Add(objChkBox)
            dgvSODetail.Columns.Add("SO", "Sale Order")
            dgvSODetail.Columns.Add("AmendmentNo", "Amendment No")
            dgvSODetail.Columns.Add("ItemCode", "Item Code")
            dgvSODetail.Columns.Add("OPENSO", "Open SO")

            dgvSODetail.Columns(enmColItem.Status).Width = 50
            dgvSODetail.Columns(enmColItem.SO).Width = 170
            dgvSODetail.Columns(enmColItem.AmendmentNo).Width = 170
            dgvSODetail.Columns(enmColItem.ItemCode).Width = 150
            dgvSODetail.Columns(enmColItem.OPENSO).Width = 100
            dgvSODetail.Columns(enmColItem.ItemCode).Visible = False

            dgvSODetail.Columns(enmColItem.Status).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvSODetail.Columns(enmColItem.SO).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvSODetail.Columns(enmColItem.AmendmentNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvSODetail.Columns(enmColItem.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvSODetail.Columns(enmColItem.OPENSO).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvSODetail.Columns(enmColItem.SO).ReadOnly = True
            dgvSODetail.Columns(enmColItem.AmendmentNo).ReadOnly = True
            dgvSODetail.Columns(enmColItem.ItemCode).ReadOnly = True
            dgvSODetail.Columns(enmColItem.OPENSO).ReadOnly = True

            dgvSODetail.Columns(enmColItem.Status).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvSODetail.Columns(enmColItem.SO).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvSODetail.Columns(enmColItem.AmendmentNo).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvSODetail.Columns(enmColItem.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvSODetail.Columns(enmColItem.OPENSO).SortMode = DataGridViewColumnSortMode.NotSortable

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub cmdcusthelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdcusthelp.Click
        Try
            Dim strCustHelp() As String
            Dim strsql As String

            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct customer_code as CustomerCode,Cust_Name as CustomerName " & _
             "from custitem_mst (nolock) a,customer_mst b  where a.Account_Code=b.Customer_Code and a.UNIT_CODE=b.UNIT_CODE and a.UNIT_CODE='" & gstrUNITID & "' and Active=1", "Customer Codes List", 1)
            If UBound(strCustHelp) <> -1 Then
                If strCustHelp(0) <> "0" Then
                    txtcustomerhelp.Text = Trim(strCustHelp(0))
                    lblCustName.Text = strCustHelp(1)
                    SetItemGridsHeader()
                    txtCustItem.Text = ""
                    lblCusttemName.Text = ""
                    txtItem.Text = ""
                    lblItemName.Text = ""
                Else
                    txtcustomerhelp.Text = ""
                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Try
            RefreshScreen()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub


    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub



    Private Sub cmdCustItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCustItem.Click
        Try
            Dim strItemHelp() As String
            Dim strsql As String
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            strItemHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct Cust_Drgno as CustDrgno,Drg_Desc as CustDrgDesc from custitem_mst (nolock) a  where  a.UNIT_CODE='" & gstrUNITID & "' and Active=1  and Account_Code='" & txtcustomerhelp.Text & "'  and Cust_Drgno in (SELECT DISTINCT  C.cust_drgno as Custdrgno from custitem_mst C with (nolock) where C.active = 1  and  unit_code='" & gstrUNITID & "' and C.ACCOUNT_CODE ='" & txtcustomerhelp.Text & "' GROUP BY C.Account_code, C.cust_drgno HAVING  COUNT(C.item_code) >1 )", "CustItem List", 1)
            If UBound(strItemHelp) <> -1 Then
                If strItemHelp(0) <> "0" Then
                    txtCustItem.Text = Trim(strItemHelp(0))
                    lblCusttemName.Text = Trim(strItemHelp(1))
                    txtItem.Text = ""
                    lblItemName.Text = ""
                Else
                    txtItem.Text = ""
                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdItemHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdItemHelp.Click
        Try
            Dim strItemHelp() As String
            Dim strsql As String
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf txtCustItem.Text = "" Then
                MsgBox("Select Customer Drg No. First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            strItemHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct Item_code as ItemCode,Item_Desc as ItemName from custitem_mst (nolock) a  where  a.UNIT_CODE='" & gstrUNITID & "' and Active=1  and Account_Code='" & txtcustomerhelp.Text & "'  and Cust_Drgno='" & txtCustItem.Text & "' ", "Item Codes List", 1)
            If UBound(strItemHelp) <> -1 Then
                If strItemHelp(0) <> "0" Then
                    txtItem.Text = Trim(strItemHelp(0))
                    lblItemName.Text = Trim(strItemHelp(1))
                Else
                    txtItem.Text = ""
                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveData()
    End Sub

    Private Sub SaveData()
        Try
            Dim Count As Integer = 0
            Dim strCustRef As String = ""
            Dim strAmendment As String = ""
            Dim strItemCode As String = ""
            Dim intCounter As Integer
            Dim ds As New DataSet

            If dgvSODetail.Rows.Count = 0 Then
                MsgBox("There is no Sale Order for update", MsgBoxStyle.Information, Me.Text)
                Exit Sub
            ElseIf dgvSODetail.Rows.Count > 0 Then
                For intCounter = 0 To dgvSODetail.Rows.Count - 1
                    If dgvSODetail.Rows(intCounter).Cells(enmColItem.Status).Value = True Then
                        Count += 1
                    End If
                Next
                If Count = 0 Then
                    MsgBox("There is no Sale Order for Update", MsgBoxStyle.Information, Me.Text)
                    Exit Sub
                ElseIf Count > 1 Then
                    MsgBox("Select only one row at a time", MsgBoxStyle.Information, Me.Text)
                    Exit Sub
                End If
            End If
            Dim strMessage As String = ""
            Dim strMessageString As String = ""
            Dim intDocNo As Integer = SqlConnectionclass.ExecuteScalar("SELECT ISNULL(MAX(DOC_NO),0) + 1 FROM Cust_Ord_Dtl_INTERNAL_ITEMLINKAGE ")
          
            For intCounter = 0 To dgvSODetail.Rows.Count - 1
                If dgvSODetail.Rows(intCounter).Cells(enmColItem.Status).Value = True Then
                    strCustRef = dgvSODetail.Rows(intCounter).Cells(enmColItem.SO).Value
                    strAmendment = dgvSODetail.Rows(intCounter).Cells(enmColItem.AmendmentNo).Value
                    'strItemCode = dgvSODetail.Rows(intCounter).Cells(enmColItem.ItemCode).Value
                    strItemCode = ""
                    Using sqlCmd As SqlCommand = New SqlCommand
                        With sqlCmd
                            .CommandType = CommandType.StoredProcedure
                            .CommandText = "USP_ADDITEM_SALEORDER"
                            .CommandTimeout = 0
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 8).Value = txtcustomerhelp.Text
                            .Parameters.Add("@NEWITEM", SqlDbType.VarChar, 16).Value = txtItem.Text
                            .Parameters.Add("@CUST_DRGNO ", SqlDbType.VarChar, 50).Value = txtCustItem.Text
                            .Parameters.Add("@CUST_REF ", SqlDbType.VarChar, 34).Value = strCustRef
                            .Parameters.Add("@AMENDMENT_NO ", SqlDbType.VarChar, 25).Value = strAmendment
                            .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                            .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 25).Value = gstrIpaddressWinSck
                            .Parameters.Add("@Doc_NO", SqlDbType.Float).Value = intDocNo
                            .Parameters.Add("@OLDITEM", SqlDbType.VarChar, 16).Value = strItemCode
                            .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                            SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                            strMessage = Convert.ToString(.Parameters("@MESSAGE").Value)
                            If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                                strMessage = Convert.ToString(.Parameters("@MESSAGE").Value)
                            End If
                            If strMessage <> "" Then
                                If strMessageString = "" Then
                                    strMessageString = "Below Sale Order are not updated.Already It has been already added " & vbCrLf & strMessage
                                Else
                                    strMessageString += vbCrLf & strMessage
                                End If
                            End If
                        End With
                    End Using
                End If

            Next
            If strMessageString <> "" Then
                MsgBox(strMessageString, MsgBoxStyle.Information, Me.Text)
            Else
                MsgBox("Item added through internal amendment successfully", MsgBoxStyle.Information, Me.Text)
                AutoMailerSend(intDocNo)
                ShowSalesOrderDATA()
            End If


        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

 
    Private Sub lblcustcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblcustcode.Click

    End Sub
	
	 Private Sub AutoMailerSend(ByVal DocNo As Integer)
        Dim FunReturn As Boolean = False
        Try
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_AUTOMAILER_ADDITEM_SALEORDER"
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