Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.IO
Imports ADODB



Friend Class FrmMKTTRN0138
    Inherits System.Windows.Forms.Form



    'Form Name           - frmMKTTRN0138
    'Form Description    - Pick List for Container Loading
    'Created by          - Sneha Aggarwal
    'Created Date        - 17/02/2025



    Dim mintFormIndex As Short
    Dim mintIndex As Short
    Dim frmPopulatedGrid As Boolean = False
    Dim pickListNo As String = String.Empty
    Dim strPickListNo As String
    Dim struser_Container_No As String
    Dim strinvoiceno As String



    Private Sub FrmMKTTRN0138_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            btn_save.Enabled = False
            btn_delete.Enabled = False
            cmdHelpCust.Enabled = True
            txtContainerno.Enabled = False
            ' optSelecedCustomers.Enabled = False


            btn_details.Enabled = False

            mintIndex = mdifrmMain.AddFormNameToWindowList("Pick List for Container Loading")
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
            Call FitToClient(Me, grpHeader, ctlFormHeader1, GroupBox1)
            Me.ctlFormHeader1.HeaderString = Mid(Me.ctlFormHeader1.HeaderString, InStr(1, Me.ctlFormHeader1.HeaderString(), "-") + 1, Len(Me.ctlFormHeader1.HeaderString()))

            Me.MdiParent = prjMPower.mdifrmMain
            'Call InitializeControls()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub



    Private Sub cmdHelpCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelpCust.Click
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim strdocno() As String
        btn_delete.Enabled = True
        btn_save.Enabled = False
        ' btn_new.Enabled = False
        lvwCustomers.Enabled = False
        lvwCustomers.Items.Clear()
        lvwCustomerselected.Enabled = True
        txtSearchCustomer.Text = ""

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        '  strQry = "select pcd.PickList_No,user_Container_No from Temp_Picklist_Verification_HDR tmptbl INNER JOIN  Picklist_Container_Loading pcd on pcd.Customer_Code=tmptbl.CustomerCode and  pcd.invoice_no=tmptbl.inv_no inner join Customer_mst cm ON tmptbl.CustomerCode = cm.Customer_Code and tmptbl.UNIT_CODE = cm.UNIT_CODE WHERE tmptbl.UNIT_CODE = '" & gstrUNITID & "'"


        strQry = "Select distinct PickList_No,Convert(Date,ENT_DT) PickDate from Picklist_Container_Loading WHERE UNIT_CODE = '" & gstrUNITID & "'"

        strdocno = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)

        If strdocno Is Nothing OrElse UBound(strdocno) = -1 Then
            MsgBox("No data returned.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Information")
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If

        strPickListNo = strdocno(0)
        struser_Container_No = strdocno(1)
        ' strinvoiceno = strdocno(2)

        txtPicklistno.Text = strPickListNo

        If Not (UBound(strdocno) = -1) Then
            If (Len(strdocno(0)) >= 1) And strdocno(0) = "0" Then
                MsgBox("No Picklist to delete.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Sub
            Else
                picklistdetailsgrid()
            End If
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)


    End Sub



    Private Function picklistdetailsgrid() As DataTable
        Dim dtData As New DataTable()
        Try

            Dim objitem As ClsResultSetDB
            Dim strSql As String
            Dim strCustomer As String
            Dim lngLoop As Integer
            Dim lngRows As Integer
            Dim intMaxCount As Short
            Dim intCount As Short
            Dim intNoRec As Short
            Dim intMaxLoop As Short
            Dim intLoopCount As Short

            ' Initialize the list view
            Me.lvwCustomerselected.Items.Clear()
            With Me.lvwCustomerselected
                .LabelEdit = False
                .CheckBoxes = False
                .View = System.Windows.Forms.View.Details
                .GridLines = True
                .Columns.Clear()
                .Columns.Insert(0, "", "Customer Code", -2)
                .Columns.Insert(1, "", "Customer Name", -2)
                .Columns.Insert(2, "", "Invoice No.", -2)
                .Columns.Insert(3, "", "Container No.", -2)
                .Columns.Item(0).Width = VB6.TwipsToPixelsX(1600)
            End With


            'strSql = "select pcd.PickList_No,user_Container_No, Pcd.Customer_Code,pcd.Customer_Name,pcd.Invoice_no from Temp_Picklist_Verification_HDR tmptbl INNER JOIN  Picklist_Container_Loading pcd on pcd.Customer_Code=tmptbl.CustomerCode and  pcd.invoice_no=tmptbl.inv_no inner join Customer_mst cm ON tmptbl.CustomerCode = cm.Customer_Code and tmptbl.UNIT_CODE = cm.UNIT_CODE WHERE tmptbl.UNIT_CODE = '" & gstrUNITID & "' and pcd.picklist_no = '" & strPickListNo & "'"
            strSql = "select pcd.PickList_No,user_Container_No, Pcd.Customer_Code,pcd.Customer_Name,pcd.Invoice_no from Picklist_Container_Loading pcd inner join Customer_mst cm on  pcd.Customer_Code=cm.Customer_Code and pcd.UNIT_CODE = cm.UNIT_CODE WHERE pcd.UNIT_CODE = '" & gstrUNITID & "' and pcd.picklist_no = '" & strPickListNo & "'"

            Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                Using SqlCmd As New SqlCommand(strSql, connection)

                    SqlCmd.Parameters.AddWithValue("@UnitCode", gstrUNITID)

                    Using da As New SqlDataAdapter(SqlCmd)

                        da.Fill(dtData)
                    End Using
                    SqlCmd.Dispose()
                End Using

            End Using

            With Me.lvwCustomerselected
                ' .Items.Clear()
                For Each row As DataRow In dtData.Rows
                    With Me.lvwCustomerselected.Items.Add(row("Customer_Code").ToString())
                        .SubItems.Add(row("Customer_Name").ToString())
                        .SubItems.Add(row("Invoice_no").ToString())
                        .SubItems.Add(row("user_Container_No").ToString())
                    End With
                Next

                Me.lvwCustomerselected.Columns.Item(0).Width = 100
                Me.lvwCustomerselected.Columns.Item(1).Width = 200
                Me.lvwCustomerselected.Columns.Item(2).Width = 200
                Me.lvwCustomerselected.Columns.Item(3).Width = 200
            End With

        Catch ex As Exception
            RaiseException(ex)
        Finally

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Function


    Private Sub Deleterecord()
        Dim strSql As String
        ' Dim strSql1 As String
        Dim selectedCustomers As New List(Of String)
        Dim selectedCustomerNames As New List(Of String)
        Dim selectedInvoiceNos As New List(Of String)
        Dim selectedContainerNos As New List(Of String)
        Dim strPickListNo As String = ""
        Dim picklistExists As Boolean = False

        Try

            For Each item As ListViewItem In lvwCustomerselected.Items
                If item.Checked Then
                    selectedCustomers.Add(item.Text)  ' Customer Code
                    selectedCustomerNames.Add(item.SubItems(1).Text)  ' Customer Name
                    selectedInvoiceNos.Add(item.SubItems(2).Text)  ' Invoice No
                    selectedContainerNos.Add(item.SubItems(3).Text)  ' Container No
                End If
            Next

            'If selectedCustomers.Count = 0 Then
            '    Me.lvwCustomerselected.Columns.Clear()
            '    Me.lvwCustomerselected.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            '    MsgBox("Select at least one Detail", MsgBoxStyle.Information, "eMPro")
            '    Exit Sub
            'End If

            'change
          
            ' Check if picklist already exists in Temp_Picklist_Verification_HDR directly
            Dim checkQuery As String = "SELECT COUNT(*) FROM Temp_Picklist_Verification_HDR WHERE picklist_no = @PicklistNo"

            Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                Using SqlCmd As New SqlCommand(checkQuery, connection)
                    SqlCmd.Parameters.AddWithValue("@PicklistNo", txtPicklistno.Text) ' Pass the picklist number to the query


                    Dim result As Integer = Convert.ToInt32(SqlCmd.ExecuteScalar())
                    If result > 0 Then
                        picklistExists = True
                    End If
                End Using
            End Using


            If picklistExists Then
                If MsgBox("The picklist is already scanned. It can not be delete", vbOKOnly + vbInformation, ResolveResString(100)) Then
                    txtPicklistno.Text = ""

                    Return

                End If
            End If


            'For i As Integer = 0 To selectedCustomers.Count - 1
            ''strSql = "Delete from Temp_Picklist_Verification_HDR where PickList_No = '" & strPickListNo & "'"
            ''strSql1 = "Delete from Temp_Picklist_Verification_dtl where PickList_No = '" & strPickListNo & "'"

            ''strSql = "Delete from Picklist_Container_Loading where PickList_No = '" & strPickListNo & "' and Invoice_no= '" & selectedInvoiceNos(i) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            strSql = "Delete from Picklist_Container_Loading where PickList_No = '" & txtPicklistno.Text & "' and UNIT_CODE = '" & gstrUNITID & "'"


            Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                Using cmd As New SqlCommand(strSql, connection)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            ''Using connection As SqlConnection = SqlConnectionclass.GetConnection()
            ''    Using cmd As New SqlCommand(strSql1, connection)
            ''        cmd.ExecuteNonQuery()
            ''    End Using
            ''End Using
            ' Next

            MsgBox("Delete successfully.", vbOKOnly + vbInformation, ResolveResString(100))
            txtPicklistno.Text = ""

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub


    Private Function populateCustomerlist() As DataTable

        Try
            Dim dtData As New DataTable()
            Dim dsData As New DataSet()
            Dim objitem As ClsResultSetDB
            Dim strSql As String
            Dim strCustomer As String
            Dim lngLoop As Integer
            Dim lngRows As Integer
            Dim intMaxCount As Short
            Dim intCount As Short
            Dim intNoRec As Short
            Dim intMaxLoop As Short
            Dim intLoopCount As Short

            ' Initialize the list view
            Me.lvwCustomerselected.Items.Clear()
            With Me.lvwCustomerselected
                .LabelEdit = False
                .CheckBoxes = True
                .View = System.Windows.Forms.View.Details
                .GridLines = True
                .Columns.Clear()
                .Columns.Insert(0, "", "Customer Code", -2)
                .Columns.Insert(1, "", "Customer Name", -2)
                .Columns.Insert(2, "", "Invoice No.", -2)
                .Columns.Insert(3, "", "Container No.", -2)
                .Columns.Item(0).Width = VB6.TwipsToPixelsX(1600)
            End With

            intMaxCount = Me.lvwCustomers.Items.Count
            For intCount = 0 To intMaxCount - 1
                If lvwCustomers.Items.Item(intCount).Checked = True Then
                    intNoRec = intNoRec + 1
                End If
            Next

            ' If no customers are selected
            If intNoRec <= 0 Then
                Me.lvwCustomerselected.Columns.Clear()
                Me.lvwCustomerselected.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                MsgBox("Select at least one Customer", MsgBoxStyle.Information, "eMPro")
                Me.lvwCustomers.Focus()
                Exit Function
            End If

            If intNoRec > 0 Then

                intMaxLoop = lvwCustomers.Items.Count
                For intLoopCount = 0 To intMaxLoop - 1
                    If lvwCustomers.Items.Item(intLoopCount).Checked = True Then
                        If Len(Trim(strCustomer)) = 0 Then
                            strCustomer = "" & lvwCustomers.Items.Item(intLoopCount).Text & "@"
                        Else
                            strCustomer = strCustomer & "" & lvwCustomers.Items.Item(intLoopCount).Text & "@"
                        End If
                    End If
                Next


                Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                    Using sqlcmd As New SqlCommand("usp_populatecustomerlist", connection)
                        sqlcmd.CommandType = CommandType.StoredProcedure
                        sqlcmd.CommandTimeout = 0
                        sqlcmd.Parameters.AddWithValue("@unit_code", gstrUNITID)
                        sqlcmd.Parameters.AddWithValue("@customer", strCustomer)


                        Dim da As New SqlDataAdapter(sqlcmd)
                        da.Fill(dsData)

                        If dsData.Tables.Count > 0 Then
                            dtData = dsData.Tables(0)
                        End If
                    End Using
                End Using


            End If

            With Me.lvwCustomerselected
                .Items.Clear()
                For Each row As DataRow In dtData.Rows
                    With Me.lvwCustomerselected.Items.Add(row("Account_code").ToString())
                        .SubItems.Add(row("CUST_NAME").ToString())
                        .SubItems.Add(row("Doc_No").ToString())
                        .SubItems.Add(row("Container_No").ToString())
                    End With
                Next

                Me.lvwCustomerselected.Columns.Item(0).Width = 100
                Me.lvwCustomerselected.Columns.Item(1).Width = 200
                Me.lvwCustomerselected.Columns.Item(2).Width = 200
                Me.lvwCustomerselected.Columns.Item(3).Width = 200
            End With


        Catch ex As Exception
            RaiseException(ex)
        Finally

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Function


    Private Sub btn_details_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_details.Click
        btn_save.Enabled = True
        btn_new.Enabled = False

        Try
            If Not frmPopulatedGrid Then
                Call populateCustomerlist()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub


    'Private Sub Customerlist()
    '    Dim dt As New DataTable()
    '    Dim strSql As String

    '    Try

    '        Me.lvwCustomers.Items.Clear()

    '        With Me.lvwCustomers
    '            .LabelEdit = False
    '            .CheckBoxes = True
    '            .View = System.Windows.Forms.View.Details
    '            .Columns.Clear()
    '            .Columns.Insert(0, "", "Customer Code", -2)
    '            .Columns.Insert(1, "", "Customer Name", -2)
    '            .Columns.Item(0).Width = VB6.TwipsToPixelsX(1400)
    '        End With


    '        strSql = "SELECT DISTINCT Customer_mst.customer_code, Customer_mst.Cust_Name " & _
    '                 "FROM Customer_mst " & _
    '                 "INNER JOIN Lists ON Customer_mst.Customer_code = Lists.Key2 " & _
    '                 "AND Customer_mst.unit_code = Lists.unit_code " & _
    '                 "WHERE Lists.Key1 LIKE '%ArticleLblCust%'"

    '        Using connection As SqlConnection = SqlConnectionclass.GetConnection()
    '            Using SqlCmd As New SqlCommand(strSql, connection)
    '                SqlCmd.Parameters.AddWithValue("@UnitCode", gstrUNITID)

    '                Using da As New SqlDataAdapter(SqlCmd)
    '                    da.Fill(dt)
    '                End Using
    '            End Using
    '        End Using

    '        If dt.Rows.Count <= 0 Then
    '            MsgBox("No Customer Codes are defined", MsgBoxStyle.Information, "eMPro")
    '            Exit Sub
    '        End If

    '        With Me.lvwCustomers
    '            '  .Items.Clear()
    '            For Each row As DataRow In dt.Rows
    '                Dim item As New ListViewItem(row("Customer_Code").ToString())
    '                item.SubItems.Add(row("CUST_NAME").ToString())
    '                .Items.Add(item)
    '            Next
    '        End With

    '        Me.lvwCustomers.Columns.Item(0).Width = 100
    '        Me.lvwCustomers.Columns.Item(1).Width = 400

    '    Catch ex As Exception

    '        RaiseException(ex)
    '    Finally

    '        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Sub



    Private Sub customerFilter()
        Dim dt As New DataTable()
        Dim dt1 As New DataTable()
        Dim strSql As String
        Dim strSql1 As String

        Try

            Me.lvwCustomers.Items.Clear()

            With Me.lvwCustomers
                .LabelEdit = False
                .CheckBoxes = True
                .View = System.Windows.Forms.View.Details
                .Columns.Clear()
                .Columns.Insert(0, "", "Customer Code", -2)
                .Columns.Insert(1, "", "Customer Name", -2)
                .Columns.Item(0).Width = VB6.TwipsToPixelsX(1400)
            End With

            If optSearchLCustCode.Checked Then
                strSql = "SELECT DISTINCT Customer_mst.customer_code, Customer_mst.Cust_Name " & _
                         "FROM Customer_mst " & _
                         "INNER JOIN Lists ON Customer_mst.Customer_code = Lists.Key2 " & _
                         "AND Customer_mst.unit_code = Lists.unit_code " & _
                         "WHERE Lists.Key1 LIKE '%ArticleLblCust%' AND Customer_mst.unit_code = @UnitCode and customer_code LIKE '%' + @custCode + '%' "

                Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                    Using SqlCmd As New SqlCommand(strSql, connection)
                        SqlCmd.Parameters.AddWithValue("@UnitCode", gstrUNITID)
                        SqlCmd.Parameters.AddWithValue("@custCode", txtSearchCustomer.Text.ToString())

                        Using da As New SqlDataAdapter(SqlCmd)
                            da.Fill(dt)
                        End Using
                    End Using
                End Using

                If dt.Rows.Count > 0 Then

                    PopulateListView(dt)
                Else

                    MsgBox("Please type a valid customer code or name.", MsgBoxStyle.Information, "eMPro")
                End If
            End If

            If optSearchCustName.Checked Then
                strSql1 = "SELECT DISTINCT Customer_mst.customer_code, Customer_mst.Cust_Name " & _
                          "FROM Customer_mst " & _
                          "INNER JOIN Lists ON Customer_mst.Customer_code = Lists.Key2 " & _
                          "AND Customer_mst.unit_code = Lists.unit_code " & _
                          "WHERE Lists.Key1 LIKE '%ArticleLblCust%' and cust_name LIKE '%' + @CustName + '%' " & _
                          "AND Customer_mst.unit_code = @UnitCode"

                Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                    Using SqlCmd As New SqlCommand(strSql1, connection)
                        SqlCmd.Parameters.AddWithValue("@UnitCode", gstrUNITID)
                        SqlCmd.Parameters.AddWithValue("@CustName", txtSearchCustomer.Text.ToString())

                        Using da As New SqlDataAdapter(SqlCmd)
                            da.Fill(dt1)
                        End Using
                    End Using
                End Using

                If dt1.Rows.Count > 0 Then

                    PopulateListView(dt1)
                Else

                    MsgBox("No customer found with the given customer name.", MsgBoxStyle.Information, "eMPro")
                End If
            End If


        Catch ex As Exception

            RaiseException(ex)
        Finally

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub PopulateListView(ByVal dt As DataTable)
        With Me.lvwCustomers

            .Items.Clear()

            For Each row As DataRow In dt.Rows
                Dim item As New ListViewItem(row("Customer_Code").ToString())
                item.SubItems.Add(row("CUST_NAME").ToString())
                .Items.Add(item)
            Next
        End With

        Me.lvwCustomers.Columns.Item(0).Width = 100
        Me.lvwCustomers.Columns.Item(1).Width = 400
    End Sub



    Private Sub btn_Customer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.lvwCustomers.Enabled = True
            Me.lvwCustomers.View = System.Windows.Forms.View.Details
            Me.lvwCustomers.CheckBoxes = True
            Me.lvwCustomers.GridLines = True
            Me.lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            btn_delete.Enabled = False
            frmPopulatedGrid = True
            txtPicklistno.Text = ""
            lvwCustomerselected.Enabled = True
            txtSearchCustomer.Enabled = True
            btn_save.Enabled = True
            optSearchLCustCode.Enabled = True
            optSearchCustName.Enabled = True
            btn_details.Enabled = True

            If txtSearchCustomer.Text = "" Then
                Call Customerlist()
            Else
                'change
                'lvwCustomerselected.Items.Clear()
                'Call customerFilter()
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            frmPopulatedGrid = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub


    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        Try
            If MessageBox.Show("Are you sure you want to Close?", "Confirmation", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            Me.Dispose()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub


    Private Sub btn_new_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_new.Click
        Try
            'btn_Customer.Enabled = True
            btn_save.Enabled = True
            btn_delete.Enabled = True
            cmdHelpCust.Enabled = False
            txtContainerno.Enabled = True
            btn_delete.Enabled = False
            txtPicklistno.Text = ""
            txtSearchCustomer.Enabled = True
            btn_details.Enabled = True
            txtSearchCustomer.Focus()
            lvwCustomers.Enabled = True
            'optSelecedCustomers.Enabled = True
            btn_details.Enabled = True
            Refresh()


            lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lvwCustomers.Enabled = True : lvwCustomers.Items.Clear()
            Call Customerlist()
          
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub


    Private Function GeneratePickList() As Boolean
        Dim pickListNo As String = String.Empty
        Dim todayDatePrefix As String = String.Empty

        Try
            ' Get current date from SQL Server
            Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                Using cmdDate As New SqlCommand("SELECT GETDATE()", connection)
                    Dim result = cmdDate.ExecuteScalar()

                    If result IsNot Nothing AndAlso IsDate(result) Then
                        todayDatePrefix = Convert.ToDateTime(result).ToString("ddMMyyyy")
                    Else
                        MessageBox.Show("Unable to get server date.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                End Using

                ' Get max picklist number for today
                Using cmd As New SqlCommand("SELECT MAX(PickList_No) FROM Picklist_Container_Loading WHERE PickList_No LIKE @DatePrefix + '%'", connection)
                    cmd.Parameters.AddWithValue("@DatePrefix", todayDatePrefix)

                    Dim result = cmd.ExecuteScalar()
                    Dim sequenceNumber As Integer = 1

                    If result IsNot DBNull.Value AndAlso result IsNot Nothing Then
                        Dim lastPickList As String = result.ToString()
                        Dim lastSequencePart As String = lastPickList.Substring(8)

                        If IsNumeric(lastSequencePart) Then
                            sequenceNumber = Convert.ToInt32(lastSequencePart) + 1

                            If sequenceNumber > 999 Then
                                MessageBox.Show("Pick list number has reached the maximum for today.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Return False
                            End If
                        End If
                    End If

                    pickListNo = todayDatePrefix & sequenceNumber.ToString("D4")
                End Using
            End Using

            txtPicklistno.Text = pickListNo
            Return True

        Catch ex As Exception
            RaiseException(ex)
            Return False

        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function


    Private Sub recordsave()
        Dim strSql As String
        Dim selectedCustomers As New List(Of String)
        Dim selectedCustomerNames As New List(Of String)
        Dim selectedInvoiceNos As New List(Of String)
        Dim selectedContainerNos As New List(Of String)

        Try

            For Each item As ListViewItem In lvwCustomerselected.Items
                If item.Checked Then

                    selectedCustomers.Add(item.Text)  ' Customer Code
                    selectedCustomerNames.Add(item.SubItems(1).Text)  ' Customer Name
                    selectedInvoiceNos.Add(item.SubItems(2).Text)  ' Invoice No
                    selectedContainerNos.Add(item.SubItems(3).Text)  ' Container No

                End If
            Next

            If selectedCustomers.Count = 0 Then
                Me.lvwCustomerselected.Columns.Clear()
                Me.lvwCustomerselected.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                MsgBox("Select at least one Detail", MsgBoxStyle.Information, "eMPro")
                Exit Sub
            End If

            ' Loop through selected items and insert them individually.
            For i As Integer = 0 To selectedCustomers.Count - 1


                'If selectedContainerNos(i).Length >= 35 Then 
                '    MsgBox("Container No exceeds maximum length.", MsgBoxStyle.Information, "eMPro")
                '    Exit Sub
                'End If

                strSql = "INSERT INTO Picklist_Container_Loading (Customer_code, Customer_Name, Invoice_No, Container_No, PickList_No, Ent_dt, Upd_dt, User_Container_no, Upd_Userid, Unit_code) " & _
                         "VALUES (@Customer_code, @Customer_Name, @Invoice_No, @Container_No, @PickList_No, GETDATE(), GETDATE(), @User_Container_no, @Upd_Userid, @Unit_code)"


                Using connection As SqlConnection = SqlConnectionclass.GetConnection()
                    Using cmd As New SqlCommand(strSql, connection)

                        cmd.Parameters.AddWithValue("@Customer_code", selectedCustomers(i))
                        cmd.Parameters.AddWithValue("@Customer_Name", selectedCustomerNames(i))
                        cmd.Parameters.AddWithValue("@Invoice_No", selectedInvoiceNos(i))
                        cmd.Parameters.AddWithValue("@Container_No", selectedContainerNos(i))
                        cmd.Parameters.AddWithValue("@PickList_No", txtPicklistno.Text.ToString())
                        cmd.Parameters.AddWithValue("@User_Container_no", txtContainerno.Text.ToString())
                        cmd.Parameters.AddWithValue("@Upd_Userid", Trim(mP_User))
                        cmd.Parameters.AddWithValue("@Unit_code", gstrUNITID)

                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            Next

            MsgBox("Save successfully.", vbOKOnly + vbInformation, ResolveResString(100))
            'txtPicklistno.Text = ""
            cmdHelpCust.Enabled = True

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub lvwCustomerselected_ItemChecked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwCustomerselected.ItemChecked

    End Sub


    Private Sub Refresh()
        Try
            txtContainerno.Text = ""
            txtSearchCustomer.Text = ""
            lvwCustomers.Items.Clear()
            lvwCustomerselected.Items.Clear()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub


    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Try

            Me.lvwCustomers.Enabled = True
            Me.lvwCustomers.View = System.Windows.Forms.View.Details
            Me.lvwCustomers.CheckBoxes = True
            Me.lvwCustomers.GridLines = True
            Me.lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

            If txtContainerno.Text = "" Then
                MsgBox("Please Type Container No.", vbOKOnly + vbInformation, ResolveResString(100))
                txtContainerno.Focus()
                Return
            ElseIf txtContainerno.Text.Length > 35 Then
                MsgBox("Invalid Container No.It should be atleast 35 character", MsgBoxStyle.Information, "eMPro")
                txtContainerno.Focus()
                Return
            End If

            Dim customerSelected As Boolean = False
            For Each item As ListViewItem In lvwCustomerselected.Items
                If item.Checked Then
                    customerSelected = True
                    Exit For
                End If
            Next

            If Not customerSelected Then
                MsgBox("Please Select Customer details", vbOKOnly + vbInformation, ResolveResString(100))
                Return
            End If

            If MessageBox.Show("Are you sure you want to save?", "Confirmation", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                If Not GeneratePickList() Then
                    ' Picklist failed, don't continue
                    Return
                End If

                recordsave()
                Refresh()
                btn_new.Enabled = True
                btn_save.Enabled = False
                optSearchLCustCode.Enabled = False
                optSearchCustName.Enabled = False
                btn_details.Enabled = False
                'optSelecedCustomers.Enabled = False
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub


    Private Sub lvwCustomers_ItemChecked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs)
       
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        Try
            Me.lvwCustomers.Enabled = True
            Me.lvwCustomers.View = System.Windows.Forms.View.Details
            Me.lvwCustomers.CheckBoxes = True
            Me.lvwCustomers.GridLines = True
            Me.lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

            Dim customerSelected As Boolean = False
            For Each item As ListViewItem In lvwCustomerselected.Items
                If item.Checked Then
                    customerSelected = True
                    Exit For
                End If
            Next


            If MessageBox.Show("Are you sure you want to Delete?", "Confirmation", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                Deleterecord()
                Refresh()
                'optSelecedCustomers.Enabled = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub


    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        'btn_Customer.Enabled = False
        btn_new.Enabled = True
        btn_save.Enabled = False
        btn_delete.Enabled = False
        cmdHelpCust.Enabled = True
        'txtContainerno.Enabled = False
        txtContainerno.Text = ""
        txtPicklistno.Text = ""
        lvwCustomers.Items.Clear()
        lvwCustomerselected.Items.Clear()
        txtSearchCustomer.Text = ""
        optSearchLCustCode.Enabled = False
        optSearchCustName.Enabled = False
        txtSearchCustomer.Enabled = False
        btn_details.Enabled = False
        'optSelecedCustomers.Enabled = False


    End Sub

 


    Public Sub search(ByRef lvwListView As System.Windows.Forms.ListView, ByRef txtSearchBox As System.Windows.Forms.TextBox, ByRef optFistOption As System.Windows.Forms.RadioButton, ByRef optSecOption As System.Windows.Forms.RadioButton)
        On Error GoTo ErrHandler
        Dim intCounter As Short
        With lvwListView
            If optFistOption.Checked = True Then
                For intCounter = 0 To .Items.Count - 1
                    If .Items.Item(intCounter).Font.Bold = True Then
                        .Items.Item(intCounter).Font = VB6.FontChangeBold(.Items.Item(intCounter).Font, False)
                        .Refresh()
                    End If
                    If .Items.Item(intCounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(intCounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(intCounter).SubItems.Item(1).Font, False)
                        .Refresh()
                    End If
                Next
                If Len(txtSearchBox.Text) = 0 Then Exit Sub
                For intCounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(intCounter).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                        .Items.Item(intCounter).Font = VB6.FontChangeBold(.Items.Item(intCounter).Font, True)
                        Call .Items.Item(intCounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            ElseIf optSecOption.Checked Then
                For intCounter = 0 To .Items.Count - 1
                    If .Items.Item(intCounter).Font.Bold = True Then
                        .Items.Item(intCounter).Font = VB6.FontChangeBold(.Items.Item(intCounter).Font, False)
                        .Refresh()
                    End If
                    If .Items.Item(intCounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(intCounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(intCounter).SubItems.Item(1).Font, False)
                        .Refresh()
                    End If
                Next
                If Len(txtSearchBox.Text) = 0 Then Exit Sub
                For intCounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(intCounter).SubItems.Item(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                        .Items.Item(intCounter).Font = VB6.FontChangeBold(.Items.Item(intCounter).Font, True)
                        Call .Items.Item(intCounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtSearchCustomer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearchCustomer.TextChanged
        Call search(lvwCustomers, txtSearchCustomer, optSearchLCustCode, optSearchCustName)
    End Sub
    Public Sub Customerlist()
        On Error GoTo ErrHandler
        Dim objcustomers As ClsResultSetDB 'Class Object
        Dim strSQLcustomers As String 'Stores the SQL statement for getting the customers
        Dim intcustomerCount As Short 'Stores the total customer count
        Dim lngCustomerCtr As Integer
        Dim strSql As String
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        lvwCustomers.Items.Clear()
        With lvwCustomers
            .Sort()
            .LabelEdit = False
            .View = System.Windows.Forms.View.Details
            .Columns.Clear()
            .Columns.Insert(0, "", "Customer Code", -2)
            .Columns.Insert(1, "", "Customer Name", -2)
            'If optAllCustomer.Checked = True Then
            '    .Enabled = False
            'Else
            '    .Enabled = True
            'End If
        End With
        strSql = "SELECT DISTINCT Customer_mst.customer_code, Customer_mst.Cust_Name " & _
                     "FROM Customer_mst " & _
                     "INNER JOIN Lists ON Customer_mst.Customer_code = Lists.Key2 " & _
                     "AND Customer_mst.unit_code = Lists.unit_code " & _
                     "WHERE Lists.Key1 LIKE '%ArticleLblCust%' and Customer_mst.unit_code = '" & gstrUNITID & "'"
        objcustomers = New ClsResultSetDB
        With objcustomers
            'Open a recordset containing the customers
            Call .GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            'Get the record count
            intcustomerCount = .GetNoRows
            If intcustomerCount <= 0 Then
                'Show message to the user
                Call MsgBox("No Customer Codes are defined.", MsgBoxStyle.Information, "empower")
                'Set the boolean variable
                ' mblnUnload = True
                'Close and release the object
                .ResultSetClose()
                objcustomers = Nothing
                'Exit from procedure
                'optAllCustomer.Checked = True
                Exit Sub
            End If
            'Populate in list view
            With lvwCustomers
                .Items.Clear()
                objcustomers.MoveFirst()
                For lngCustomerCtr = 0 To intcustomerCount - 1
                    .Items.Insert(lngCustomerCtr, objcustomers.GetValue("Customer_code"))
                    .Items.Item(lngCustomerCtr).SubItems.Add(objcustomers.GetValue("Cust_NAme"))
                    objcustomers.MoveNext()
                Next
            End With

            Me.lvwCustomers.Columns.Item(0).Width = 100
            Me.lvwCustomers.Columns.Item(1).Width = 400
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    'Private Sub optSelecedCustomers_CheckedChanged(ByVal eventSender As System.Object, ByVal e As System.EventArgs)
    '    'If eventSender.Checked Then
    '    '    If optAllCustomer.Checked = True Then
    '    '        optSearchLCustCode.Checked = True : optSearchCustName.Checked = False : txtSearchCustomer.Text = "" : optSearchLCustCode.Enabled = False
    '    '        optSearchCustName.Enabled = False : txtSearchCustomer.Enabled = False : txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
    '    '        lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : lvwCustomers.Enabled = False : lvwCustomers.Items.Clear()
    '    '    Else
    '    '        optSearchLCustCode.Checked = False : optSearchCustName.Checked = True : txtSearchCustomer.Text = "" : optSearchLCustCode.Enabled = True
    '    '        optSearchCustName.Enabled = True : txtSearchCustomer.Enabled = True : txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
    '    '        lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lvwCustomers.Enabled = True : lvwCustomers.Items.Clear()
    '    '        Call Customerlist()
    '    '    End If
    '    'End If
    'End Sub

    Private Sub lvwCustomers_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lvwCustomers.ItemCheck
        Dim Item As System.Windows.Forms.ListViewItem = lvwCustomers.Items(EventArgs.Index)

    End Sub

    'Private Sub optSelecedCustomers_Click(ByVal eventSender As System.Object, ByVal e As System.EventArgs)
    '    btn_details.Enabled = True
    '    If eventSender.Checked Then
    '        If optAllCustomer.Checked = True Then
    '            optSearchLCustCode.Checked = True : optSearchCustName.Checked = False : txtSearchCustomer.Text = "" : optSearchLCustCode.Enabled = False
    '            optSearchCustName.Enabled = False : txtSearchCustomer.Enabled = False : txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
    '            lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : lvwCustomers.Enabled = False : lvwCustomers.Items.Clear()
    '        Else
    '            optSearchLCustCode.Checked = False : optSearchCustName.Checked = True : txtSearchCustomer.Text = "" : optSearchLCustCode.Enabled = True
    '            optSearchCustName.Enabled = True : txtSearchCustomer.Enabled = True : txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
    '            lvwCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lvwCustomers.Enabled = True : lvwCustomers.Items.Clear()
    '            Call Customerlist()
    '        End If
    '    End If
    'End Sub


End Class