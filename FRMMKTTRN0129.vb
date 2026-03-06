Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Text
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections.Generic

Friend Class FRMMKTTRN0129
    Inherits System.Windows.Forms.Form
    '=====================================================================
    'Copyright(c)   - MTSL
    'Created by     - Gaurav Kumar
    'Created Date   - 22-08-2023
    'Description    - Create Template for ASN vs GRN
    '=====================================================================
    Dim mintIndex As Short
    Dim mintRecRetrieved As Short
    Dim dt As DataTable = New DataTable()

    Private Sub FRMMKTTRN0129_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, Frame1, ctlFormHeader1, Button1)
            Me.MdiParent = mdifrmMain
            Me.Icon = mdifrmMain.Icon
            mintIndex = mdifrmMain.AddFormNameToWindowList("ASN VS GRIN Template")
            chkTempActive.Checked = True
            Button1.Visible = False
            cmbColumnType.Items.Clear()
            cmbColumnType.Items.Add("Text")
            cmbColumnType.Items.Add("Number")
            'cmbColumnType.Items.Add("Boolean")
            cmbColumnType.Items.Add("Datetime")
            cmbColumnType.SelectedIndex = -1

            dt.Clear()
            dt.Columns.Add("Column_Name", GetType(String))
            dt.Columns.Add("Column_Type", GetType(String))
            dt.Columns.Add("Column_Length", GetType(String))

            lblColumnLength.Enabled = False
            txtColumnLength.Enabled = False
            txtTemplateUpload.Enabled = False
            txtCustomerCodeUpload.Enabled = False
            txtUploadID.Enabled = False
            btnUploadExcel.Enabled = False
            DisableUploadControl()
            DisableControl()

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmbColumnType_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbColumnType.SelectionChangeCommitted
        Try
            If cmbColumnType.Text.ToString() = "Text" Then
                lblColumnLength.Enabled = True
                txtColumnLength.Enabled = True
            Else
                lblColumnLength.Enabled = False
                txtColumnLength.Enabled = False
                txtColumnLength.Text = String.Empty
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            If Len(txtColumnName.Text) = 0 Then
                MsgBox("Column Name can not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If
            If cmbColumnType.Text.Trim = "" Then
                MsgBox("Column Type can not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If

            If cmbColumnType.Text.Trim = "Text" Then
                If Len(txtColumnLength.Text) = 0 Then
                    MsgBox("Column Length can not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    Return
                End If
            End If

            Dim _txtColumnName As String = RemoveSpecialCharacters(txtColumnName.Text.Trim().Replace(" ", "_"))
            If (String.IsNullOrEmpty(_txtColumnName)) Then
                MsgBox("Column name should be have characters.", MsgBoxStyle.Information, ResolveResString(100))
                Return
            End If

            txtColumnName.Text = _txtColumnName

            If cmdDocumentNo.Enabled = True AndAlso btnAdd.Text.ToUpper() = "ADD" Then

                Dim strQry As StringBuilder = New StringBuilder()
                strQry.Append("ALTER TABLE EMPRO_DYN_")
                strQry.Append(txtTemplate_Sch.Text)
                strQry.Append(" ADD ")
                strQry.Append(txtColumnName.Text)
                Select Case cmbColumnType.Text.ToString()
                    Case "Text"
                        strQry.Append(" VARCHAR")
                    Case "Number"
                        strQry.Append(" Numeric(19,4)")
                    Case "Boolean"
                        strQry.Append(" bit")
                    Case "Datetime"
                        strQry.Append(" Datetime,")
                End Select
                If cmbColumnType.Text = "Text" Then
                    strQry.Append("(" & txtColumnLength.Text & ")")
                End If
                strQry.Append(";select column_name Column_Name,case when data_type='varchar' then 'Text' when DATA_TYPE='datetime' then 'Datetime' when DATA_TYPE='numeric' then 'Number' when DATA_TYPE='bit' then 'Boolean' else data_type end Column_Type,ISNULL(character_maximum_length,0) Column_Length from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='")
                strQry.Append("EMPRO_DYN_")
                strQry.Append(txtTemplate_Sch.Text)
                strQry.Append("' and column_name not in('Customer_Code','Upload_Type','UPLOAD_ID','unit_code','Ent_dt','Ent_Userid','Upd_dt','Upd_Userid')")
                Using dtRetLst As DataTable = SqlConnectionclass.GetDataTable(strQry.ToString())
                    If dtRetLst IsNot Nothing AndAlso dtRetLst.Rows.Count > 0 Then
                        grd_Template.DataSource = dtRetLst
                        grd_Template.AllowUserToAddRows = False
                        dt = TryCast(grd_Template.DataSource, DataTable)
                        dt.AcceptChanges()
                    End If
                End Using
            ElseIf cmdDocumentNo.Enabled = True AndAlso btnAdd.Text.ToUpper() = "UPDATE" Then
                Dim retCLN_LEN As Integer = 0
                Dim _retOuput As String
                _retOuput = SqlConnectionclass.ExecuteScalar(" SELECT DATA_TYPE+'_'+cast(ISNULL(character_maximum_length,0) as varchar(10))+'_'+(select cast(isnull(count(*),0) as varchar(20)) from EMPRO_DYN_" & txtTemplate_Sch.Text & ") from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & txtTemplate_Sch.Text & "' AND UPPER(COLUMN_NAME)=UPPER('" & txtColumnName.Text & "')")
                retCLN_LEN = _retOuput.Split("_")(1)
                If (_retOuput.Split("_")(0).ToUpper() = "VARCHAR" AndAlso cmbColumnType.Text.ToUpper() = "TEXT" AndAlso _retOuput.Split("_")(2) > 0) Then
                    If retCLN_LEN <= txtColumnLength.Text Then
                        Try
                            Using dtRetLst As DataTable = SqlConnectionclass.GetDataTable(" ALTER TABLE EMPRO_DYN_" & txtTemplate_Sch.Text & " ALTER COLUMN " & txtColumnName.Text & " VARCHAR (" & txtColumnLength.Text & ");select column_name Column_Name,case when data_type='varchar' then 'Text' when DATA_TYPE='datetime' then 'Datetime' when DATA_TYPE='numeric' then 'Number' when DATA_TYPE='bit' then 'Boolean' else data_type end Column_Type,ISNULL(character_maximum_length,0) Column_Length from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & txtTemplate_Sch.Text & "' and column_name not in('Customer_Code','Upload_Type','UPLOAD_ID','unit_code','Ent_dt','Ent_Userid','Upd_dt','Upd_Userid')")
                                If dtRetLst IsNot Nothing AndAlso dtRetLst.Rows.Count > 0 Then
                                    grd_Template.DataSource = dtRetLst
                                    grd_Template.AllowUserToAddRows = False
                                    dt = TryCast(grd_Template.DataSource, DataTable)
                                    dt.AcceptChanges()
                                    btnAdd.Text = "Add"
                                End If
                            End Using
                        Catch ex As Exception
                            RaiseException(ex)
                        End Try
                    Else
                        MsgBox("You can't decrease length of column.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return
                    End If

                Else
                    Dim strQry As StringBuilder = New StringBuilder()
                    Select Case cmbColumnType.Text.ToString()
                        Case "Text"
                            strQry.Append(" VARCHAR(")
                            strQry.Append(txtColumnLength.Text)
                            strQry.Append(")")
                        Case "Number"
                            strQry.Append(" Numeric(19,4)")
                        Case "Boolean"
                            strQry.Append(" bit")
                        Case "Datetime"
                            strQry.Append(" Datetime")
                    End Select
                    Try
                        Using dtRetLst As DataTable = SqlConnectionclass.GetDataTable(" ALTER TABLE EMPRO_DYN_" & txtTemplate_Sch.Text & " ALTER COLUMN " & txtColumnName.Text & "" & strQry.ToString() & " ;select column_name Column_Name,case when data_type='varchar' then 'Text' when DATA_TYPE='datetime' then 'Datetime' when DATA_TYPE='numeric' then 'Number' when DATA_TYPE='bit' then 'Boolean' else data_type end Column_Type,ISNULL(character_maximum_length,0) Column_Length from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & txtTemplate_Sch.Text & "' and column_name not in('Customer_Code','Upload_Type','UPLOAD_ID','unit_code','Ent_dt','Ent_Userid','Upd_dt','Upd_Userid')")
                            If dtRetLst IsNot Nothing AndAlso dtRetLst.Rows.Count > 0 Then
                                grd_Template.DataSource = dtRetLst
                                grd_Template.AllowUserToAddRows = False
                                dt = TryCast(grd_Template.DataSource, DataTable)
                                dt.AcceptChanges()
                                btnAdd.Text = "Add"
                            End If
                        End Using
                    Catch ex As Exception
                        Dim _errorMsg As String = ex.Message.Replace("varchar", "Text").Replace("data type", "type").Replace("numeric", "Number")
                        _errorMsg = _errorMsg + " bcause data exits against this template in database."
                        MsgBox(_errorMsg, MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return
                    End Try
                End If
                'retCLN_LEN = SqlConnectionclass.ExecuteScalar(" SELECT ISNULL(character_maximum_length,0) from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & txtTemplate_Sch.Text & "' AND DATA_TYPE='VARCHAR' AND UPPER(COLUMN_NAME)=UPPER('" & txtColumnName.Text & "')")
                'If retCLN_LEN < txtColumnLength.Text Then
                '    Using dtRetLst As DataTable = SqlConnectionclass.GetDataTable(" ALTER TABLE EMPRO_DYN_" & txtTemplate_Sch.Text & " ALTER COLUMN " & txtColumnName.Text & " VARCHAR (" & txtColumnLength.Text & ");select column_name Column_Name,case when data_type='varchar' then 'Text' when DATA_TYPE='numeric' then 'Number' when DATA_TYPE='bit' then 'Boolean' else data_type end Column_Type,ISNULL(character_maximum_length,0) Column_Length from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & txtTemplate_Sch.Text & "' and column_name not in('Customer_Code','Upload_Type','UPLOAD_ID','unit_code','Ent_dt','Ent_Userid','Upd_dt','Upd_Userid')")
                '        If dtRetLst IsNot Nothing AndAlso dtRetLst.Rows.Count > 0 Then
                '            grd_Template.DataSource = dtRetLst
                '            grd_Template.AllowUserToAddRows = False
                '            dt = TryCast(grd_Template.DataSource, DataTable)
                '            dt.AcceptChanges()
                '            btnAdd.Text = "Add"
                '        End If
                '    End Using
                'Else
                '    MsgBox("You can't change length of column.", MsgBoxStyle.Exclamation, ResolveResString(100))
                '    Return
                'End If
            Else
                Dim row As DataRow = dt.NewRow()
                row("Column_Name") = txtColumnName.Text
                row("Column_Type") = cmbColumnType.Text
                row("Column_Length") = If(cmbColumnType.Text = "Text", txtColumnLength.Text, "0")
                dt.Rows.Add(row)
                grd_Template.DataSource = dt
                Dim DeleteButtonColumn As DataGridViewButtonColumn = New DataGridViewButtonColumn()
                DeleteButtonColumn.Name = "Delete_column"
                DeleteButtonColumn.Text = "Delete"
                DeleteButtonColumn.HeaderText = "Delete"
                DeleteButtonColumn.UseColumnTextForButtonValue = True
                Dim columnIndex As Integer = 3

                If grd_Template.Columns("Delete_column") Is Nothing Then
                    grd_Template.Columns.Insert(columnIndex, DeleteButtonColumn)
                End If
            End If
            txtColumnName.Text = String.Empty
            cmbColumnType.SelectedIndex = -1
            txtColumnLength.Text = String.Empty
            lblColumnLength.Enabled = False
            txtColumnLength.Enabled = False
            txtColumnName.Focus()

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub grd_Template_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grd_Template.CellClick
        Try
            Dim intResult As Integer = 0
            If grd_Template.Columns.Contains("Delete_column") = True Then
                If e.ColumnIndex = grd_Template.Columns("Delete_column").Index Then
                    intResult = MessageBox.Show(" Do you want to delete?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    If intResult = DialogResult.Yes Then
                        grd_Template.Rows.RemoveAt(e.RowIndex)
                        dt = TryCast(grd_Template.DataSource, DataTable)
                        dt.AcceptChanges()
                        grd_Template.DataSource = dt
                        grd_Template.AllowUserToAddRows = False
                    End If
                End If
            End If

            If grd_Template.Columns.Contains("Edit_column") = True Then
                If e.ColumnIndex = grd_Template.Columns("Edit_column").Index Then
                    If e.RowIndex > -1 Then
                        txtColumnName.Text = grd_Template.Rows(e.RowIndex).Cells("Column_Name").Value.ToString()
                        cmbColumnType.SelectedItem = grd_Template.Rows(e.RowIndex).Cells("Column_Type").Value.ToString()
                        If grd_Template.Rows(e.RowIndex).Cells("Column_Type").Value.ToString().ToUpper() = "TEXT" Then
                            txtColumnLength.Text = grd_Template.Rows(e.RowIndex).Cells("Column_Length").Value.ToString()
                            txtColumnLength.Enabled = True
                            lblColumnLength.Enabled = True
                            btnAdd.Text = "Update"
                            'cmbColumnType.Enabled = False
                        Else
                            txtColumnLength.Text = String.Empty
                            txtColumnLength.Enabled = False
                            lblColumnLength.Enabled = False
                            btnAdd.Text = "Update"
                        End If
                        'Else
                        '    txtColumnName.Text = String.Empty
                        '    cmbColumnType.SelectedIndex = -1
                        '    txtColumnLength.Enabled = False
                        '    txtColumnLength.Text = String.Empty
                        '    btnAdd.Text = "Add"
                        '    cmbColumnType.Enabled = True
                        'End If
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub grd_Template_DataBindingComplete(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles grd_Template.DataBindingComplete
        Try
            For Each row As DataGridViewRow In grd_Template.Rows
                grd_Template.Rows(row.Index).HeaderCell.Value = String.Format("{0}  ", row.Index + 1).ToString()
                grd_Template.RowHeadersWidth = 80
            Next
            If grd_Template.Columns.Contains("Delete_column") = True Then
                grd_Template.Columns("Delete_column").Width = 150
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtColumnLength_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtColumnLength.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    'Private Sub txtColumnName_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtColumnName.KeyPress
    '    If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsLetter(e.KeyChar) Then
    '        e.Handled = True
    '    End If
    'End Sub

    Private Sub btnCreateTemplate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateTemplate.Click
        Try
            Dim intResult As Integer = 0
            Dim cmd As SqlCommand
            Dim SqlAdp As SqlDataAdapter
            Dim dtReturn As DataTable
            Dim templateName As String

            intResult = MessageBox.Show(" Do you want to create template?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            If intResult = DialogResult.Yes Then
                If grd_Template.Rows.Count > 0 Then
                    If Len(txtTemplateName.Text) = 0 Then
                        MsgBox("Template Name can not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return
                    Else
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                        Dim dtTable As DataTable = New DataTable()
                        dtTable = TryCast(grd_Template.DataSource, DataTable)
                        Dim strQuery As StringBuilder = New StringBuilder()
                        strQuery.Append("Create Table ")
                        templateName = "EMPRO_DYN_" + txtTemplateName.Text
                        strQuery.Append(templateName)
                        strQuery.Append("( ")
                        For i As Int32 = 0 To dtTable.Rows.Count - 1
                            strQuery.Append(dtTable.Rows(i)("Column_Name").ToString().Replace(" ", "_") + " ")
                            If dtTable.Rows(i)("Column_Type").ToString() = "Text" Then
                                strQuery.Append("VARCHAR")
                            Else
                                Select Case dtTable.Rows(i)("Column_Type").ToString()
                                    Case "Number"
                                        strQuery.Append("Numeric(19,4),")
                                    Case "Boolean"
                                        strQuery.Append("bit,")
                                    Case "Datetime"
                                        strQuery.Append("Datetime,")
                                End Select
                            End If
                            If dtTable.Rows(i)("Column_Type").ToString() = "Text" Then
                                strQuery.Append("(" & dtTable.Rows(i)("Column_Length").ToString() & "),")
                            End If
                        Next
                        strQuery.Append("Customer_Code varchar(8),Upload_Type varchar(8),UPLOAD_ID varchar(17),unit_code varchar(10),Ent_dt datetime,Ent_Userid varchar(16),Upd_dt datetime,Upd_Userid varchar(16)")
                        strQuery.Append(")")
                        cmd = New SqlCommand
                        With cmd
                            .CommandText = "USP_GRN_TEMPLATE_CREATION"
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .Connection = SqlConnectionclass.GetConnection()
                            .Parameters.Clear()
                            .Parameters.AddWithValue("@TEMPLATE_NAME", SqlDbType.VarChar).Value = txtTemplateName.Text.Trim()
                            .Parameters.AddWithValue("@TEMPLATE_TABLE", SqlDbType.VarChar).Value = strQuery.ToString()
                            .Parameters.AddWithValue("@ACTIVE", chkTempActive.Checked)
                            .Parameters.AddWithValue("@USERID", mP_User)
                            .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                            .Parameters.AddWithValue("@TEMPLATE_MAPPING_TABLE", templateName)
                            SqlAdp = New SqlDataAdapter
                            SqlAdp.SelectCommand = cmd
                            dtReturn = New DataTable()
                            SqlAdp.Fill(dtReturn)
                            .Dispose()
                            If dtReturn IsNot Nothing AndAlso dtReturn.Rows.Count > 0 Then
                                If Convert.ToString(dtReturn.Rows(0).Item("MSG")) = "SUCCESS" Then
                                    dt.Clear()
                                    grd_Template.DataSource = dt
                                    'grd_Template.Columns.Clear()
                                    txtTemplateName.Text = String.Empty
                                    MsgBox("Template is created successfully.", MsgBoxStyle.Information, ResolveResString(100))
                                Else
                                    MsgBox(dtReturn.Rows(0).Item("MSG"), MsgBoxStyle.Exclamation, ResolveResString(100))
                                End If

                            End If
                        End With
                    End If
                Else
                    MsgBox("Grid can't be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    Return
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub txtTemplateName_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTemplateName.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsLetter(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Dim intResult As Integer = 0
        intResult = MsgBox("Do You Want To Close this Screen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100))
        If intResult = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub cmdDocumentNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDocumentNo.Click
        Dim strSql As String = String.Empty
        Dim strHelp() As String = Nothing
        Try

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            strSql = "Select TEMPLATE_ID,TEMPLATE_NAME,TEMPLATE_DT from GRN_TEMPLATE_MST WHERE ACTIVE=1 AND UNIT_CODE='" & gstrUNITID & "' ORDER BY TEMPLATE_DT"
            strHelp = BtnHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, ResolveResString(100), 0)
            If UBound(strHelp) = -1 Then Exit Sub
            If strHelp(0) <> "0" Then
                txtTemplate_Sch.Text = strHelp(1).Trim
                txtTemplate_Sch.Enabled = False
                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandText = "select column_name Column_Name,case when data_type='varchar' then 'Text' when DATA_TYPE='datetime' then 'Datetime'  when DATA_TYPE='numeric' then 'Number' when DATA_TYPE='bit' then 'Boolean' else data_type end Column_Type,ISNULL(character_maximum_length,0) Column_Length from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & txtTemplate_Sch.Text & "' and column_name not in('Customer_Code','Upload_Type','UPLOAD_ID','unit_code','Ent_dt','Ent_Userid','Upd_dt','Upd_Userid')"
                        .CommandTimeout = 0
                        .CommandType = CommandType.Text
                        Using dtRetLst As DataTable = SqlConnectionclass.GetDataTable(sqlCmd)
                            If dtRetLst IsNot Nothing AndAlso dtRetLst.Rows.Count > 0 Then
                                grd_Template.DataSource = dtRetLst
                                Dim EditButtonColumn As DataGridViewButtonColumn = New DataGridViewButtonColumn()
                                EditButtonColumn.Name = "Edit_column"
                                EditButtonColumn.Text = "Edit"
                                EditButtonColumn.HeaderText = "Edit"
                                EditButtonColumn.UseColumnTextForButtonValue = True
                                Dim columnIndex As Integer = 3

                                If grd_Template.Columns("Edit_column") Is Nothing Then
                                    grd_Template.Columns.Insert(columnIndex, EditButtonColumn)
                                End If
                                grd_Template.AllowUserToAddRows = False
                                dt = TryCast(grd_Template.DataSource, DataTable)
                                dt.AcceptChanges()
                                txtColumnName.Enabled = True
                                cmbColumnType.Enabled = True
                                btnAdd.Enabled = True
                            Else
                                txtColumnName.Enabled = False
                                cmbColumnType.Enabled = False
                                btnAdd.Enabled = False
                                MsgBox("No record found !", MsgBoxStyle.Information, ResolveResString(100))
                                Return
                            End If
                            txtColumnName.Text = String.Empty
                            txtColumnLength.Text = String.Empty
                            lblColumnLength.Enabled = False
                            txtColumnLength.Enabled = False
                            cmbColumnType.SelectedIndex = -1
                            btnAdd.Text = "Add"
                        End Using
                    End With
                End Using
            Else
                MsgBox("No Record found !", MsgBoxStyle.Information, ResolveResString(100))
                txtTemplate_Sch.Text = String.Empty
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        txtTemplate_Sch.Enabled = False
        cmdDocumentNo.Enabled = False
        txtTemplate_Sch.Text = String.Empty
        EnableControl()
        dt.Clear()
        grd_Template.DataSource = dt
        grd_Template.Columns.Clear()
        grd_Template.DataSource = Nothing
        grd_Template.AllowUserToAddRows = False
        btnAdd.Text = "Add"
        txtColumnName.Text = String.Empty
        txtColumnLength.Text = String.Empty
        cmbColumnType.SelectedIndex = -1
        txtColumnLength.Enabled = False
    End Sub

    Private Sub DisableControl()
        Try

            txtColumnName.Enabled = False
            cmbColumnType.Enabled = False
            btnAdd.Enabled = False
            grd_Template.DataSource = Nothing
            txtTemplateName.Enabled = False
            chkTempActive.Enabled = False
            btnCreateTemplate.Enabled = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub EnableControl()
        Try
            txtColumnName.Enabled = True
            cmbColumnType.Enabled = True
            btnAdd.Enabled = True
            grd_Template.DataSource = Nothing
            txtTemplateName.Enabled = True
            chkTempActive.Enabled = True
            btnCreateTemplate.Enabled = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub


    Private Sub btnFileUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFileUpload.Click
        Try
            OpenFileDialog1.InitialDirectory = gstrLocalCDrive
            OpenFileDialog1.Filter = "Microsoft Excel File (*.xls)|*.xls"
            OpenFileDialog1.ShowDialog()
            If OpenFileDialog1.FileName.ToLower() <> "openfiledialog1" Then
                Me.txtFileUpload.Text = OpenFileDialog1.FileName
                Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & OpenFileDialog1.FileName & "; Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'"
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim oledbConn As OleDbConnection = New OleDbConnection(connString)
                oledbConn.Open()
                Dim dtExcelName As DataTable = oledbConn.GetSchema("Tables")
                Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM [" & dtExcelName.Rows(0)("TABLE_NAME") & "]", oledbConn)
                Dim oleda As OleDbDataAdapter = New OleDbDataAdapter()
                oleda.SelectCommand = cmd
                Dim ds As DataSet = New DataSet()
                oleda.Fill(ds)
                If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                    Dim dtExcel As DataTable
                    dtExcel = ds.Tables(0)
                    If dtExcel IsNot Nothing AndAlso dtExcel.Rows.Count > 0 Then
                        For Each column As DataColumn In dtExcel.Columns
                            Dim _ColName As String = RemoveSpecialCharacters(column.ColumnName.Trim().Replace(" ", "_"))
                            If (String.IsNullOrEmpty(_ColName)) Then
                                MsgBox("Excel column name [" & column.ColumnName & "] should be have characters.", MsgBoxStyle.Information, ResolveResString(100))
                                dtExcel.Reset()
                                txtFileUpload.Text = String.Empty
                                Return
                            End If
                            column.ColumnName = _ColName
                        Next
                        dtExcel = dtExcel.Rows.Cast(Of DataRow)().Where(Function(row) Not row.ItemArray.All(Function(f) TypeOf f Is DBNull)).CopyToDataTable()
                        grdExcelData.DataSource = dtExcel
                        grd_Template.AllowUserToAddRows = False
                        grdExcelData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Public Shared Function RemoveSpecialCharacters(ByVal str As String) As String
        Dim sb As StringBuilder = New StringBuilder()
        For Each c As Char In str
            If (c >= "0"c AndAlso c <= "9"c) OrElse (c >= "A"c AndAlso c <= "Z"c) OrElse (c >= "a"c AndAlso c <= "z"c) OrElse c = "_"c Then
                sb.Append(c)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Sub btnTemplateMatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTemplateMatch.Click
        Dim strSql As String = String.Empty
        Dim strHelp() As String = Nothing
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            strSql = "Select TEMPLATE_ID,TEMPLATE_NAME,TEMPLATE_DT from GRN_TEMPLATE_MST WHERE ACTIVE=1 AND UNIT_CODE='" & gstrUNITID & "' ORDER BY TEMPLATE_DT"
            strHelp = BtnHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, ResolveResString(100), 0)
            If UBound(strHelp) = -1 Then Exit Sub
            If strHelp(0) <> "0" Then
                txtTemplateUpload.Text = strHelp(1).Trim
                txtTemplateUpload.Enabled = False
                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandText = "select column_name Column_Name,case when data_type='varchar' then 'Text' when DATA_TYPE='numeric' then 'Number' when DATA_TYPE='bit' then 'Boolean' else data_type end Column_Type,ISNULL(character_maximum_length,0) Column_Length from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & strHelp(1).Trim & "' and column_name not in('Customer_Code','Upload_Type','UPLOAD_ID','unit_code','Ent_dt','Ent_Userid','Upd_dt','Upd_Userid')"
                        .CommandTimeout = 0
                        .CommandType = CommandType.Text
                        Using dtRetLst As DataTable = SqlConnectionclass.GetDataTable(sqlCmd)
                            If dtRetLst IsNot Nothing AndAlso dtRetLst.Rows.Count > 0 Then
                                grdTempMatch.DataSource = dtRetLst
                                grdTempMatch.AllowUserToAddRows = False
                            Else
                                MsgBox("No record found !", MsgBoxStyle.Information, ResolveResString(100))
                                Return
                            End If
                        End Using
                    End With
                End Using
            Else
                MsgBox("No Record found !", MsgBoxStyle.Information, ResolveResString(100))
                txtTemplate_Sch.Text = String.Empty
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub


    Private Sub btnUploadExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUploadExcel.Click
        Try
            Dim dtUploadExcel As DataTable = Nothing
            Dim dtColMatch As DataTable = Nothing
            Dim tableColumnNames As List(Of String) = Nothing
            Dim schemaColumnNames As List(Of String) = Nothing
            Dim unmatchedColumnNames As List(Of String) = Nothing
            Dim result As String = Nothing
            Dim rbPaymentValue As String = Nothing

            If Len(txtCustomerCodeUpload.Text) = 0 Then
                MsgBox("Customer code can not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If

            If Len(txtTemplateUpload.Text) = 0 Then
                MsgBox("Template can not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If
            If rbPayment.Checked = False AndAlso rbGRIN.Checked = False Then
                MsgBox("Upload type can not be uncheck.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If

            If grdExcelData IsNot Nothing AndAlso grdExcelData.Rows.Count > 0 Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                dtUploadExcel = TryCast(grdExcelData.DataSource, DataTable)
                If dtUploadExcel IsNot Nothing AndAlso dtUploadExcel.Rows.Count > 0 Then
                    dtColMatch = SqlConnectionclass.GetDataTable("Select * from EMPRO_DYN_" & txtTemplateUpload.Text & " WHERE 1>2")
                    dtColMatch.Columns.Remove("UPLOAD_ID")
                    dtColMatch.Columns.Remove("unit_code")
                    dtColMatch.Columns.Remove("Ent_dt")
                    dtColMatch.Columns.Remove("Ent_Userid")
                    dtColMatch.Columns.Remove("Upd_dt")
                    dtColMatch.Columns.Remove("Upd_Userid")

                    tableColumnNames = dtUploadExcel.Columns.Cast(Of DataColumn)().[Select](Function(c) c.ColumnName.ToUpper()).ToList()
                    schemaColumnNames = dtColMatch.Columns.Cast(Of DataColumn)().[Select](Function(c) c.ColumnName.ToUpper()).ToList()
                    unmatchedColumnNames = (From col In tableColumnNames Where Not schemaColumnNames.Contains(col) Select col).ToList()
                    If (unmatchedColumnNames IsNot Nothing AndAlso unmatchedColumnNames.Count > 0) Then
                        result = unmatchedColumnNames.Aggregate("", Function(current, s) current + (s & ","))
                        MsgBox("File Columns [ " & result.Remove(result.Length - 1) & " ] does not match with template.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return
                    Else
                        For i As Int64 = 0 To tableColumnNames.Count - 1
                            Dim _colTblValue As DataTable = SqlConnectionclass.GetDataTable("select isnull(character_maximum_length,0) Column_Length,DATA_TYPE from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='EMPRO_DYN_" & txtTemplateUpload.Text & "' and column_name='" & tableColumnNames(i) & "'")
                            If _colTblValue IsNot Nothing AndAlso _colTblValue.Rows.Count > 0 Then
                                If _colTblValue.Rows(0)("DATA_TYPE").ToString().ToUpper() = "VARCHAR" Then
                                    Dim lstIssueRecord As List(Of String) = dtUploadExcel.AsEnumerable().[Select](Function(row) System.Convert.ToString(row(tableColumnNames(i)))).Where(Function(x) x.Length > Convert.ToInt64(_colTblValue.Rows(0)("Column_Length"))).ToList()
                                    If lstIssueRecord.Count > 0 Then
                                        'Dim source = New BindingSource(lstIssueRecord, Nothing)
                                        grdExcelData.DataSource = Nothing
                                        Dim gridData As DataTable = New DataTable()
                                        gridData.Columns.Add("" & tableColumnNames(i) & "=>Below record length is more than database column")
                                        For Each listItem As String In lstIssueRecord
                                            gridData.Rows.Add(listItem)
                                        Next
                                        Dim gridDataBinder As BindingSource = New BindingSource()
                                        gridDataBinder.DataSource = gridData
                                        grdExcelData.DataSource = gridDataBinder
                                        grdExcelData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                                        'Dim _resultMsg As String = Nothing
                                        '_resultMsg = String.Join(",", lstIssueRecord.ToArray())
                                        'MsgBox("" & tableColumnNames(i) & "=>Below record length is more than database column :- " & vbNewLine & "" & _resultMsg & "", MsgBoxStyle.Exclamation, ResolveResString(100))
                                        Return
                                    End If
                                End If
                            End If
                        Next
                        Dim _currentDatetime As DateTime = GetServerDateTime()

                        Dim Column_Customer_Code As New Data.DataColumn("Customer_Code", GetType(System.String))
                        Column_Customer_Code.DefaultValue = txtCustomerCodeUpload.Text
                        dtUploadExcel.Columns.Add(Column_Customer_Code)

                        Dim Column_Upload_Type As New Data.DataColumn("Upload_Type", GetType(System.String))
                        Dim isChecked As Boolean = rbPayment.Checked
                        If isChecked Then
                            Column_Upload_Type.DefaultValue = rbPayment.Text
                            rbPaymentValue = rbPayment.Text
                        Else
                            Column_Upload_Type.DefaultValue = rbGRIN.Text
                            rbPaymentValue = rbGRIN.Text
                        End If
                        dtUploadExcel.Columns.Add(Column_Upload_Type)

                        Dim Column_UPLOAD_ID As New Data.DataColumn("UPLOAD_ID", GetType(System.String))
                        Column_UPLOAD_ID.DefaultValue = _currentDatetime.ToString("ddMMyyyyhhmmss")
                        dtUploadExcel.Columns.Add(Column_UPLOAD_ID)

                        Dim Column_unit_code As New Data.DataColumn("unit_code", GetType(System.String))
                        Column_unit_code.DefaultValue = gstrUNITID
                        dtUploadExcel.Columns.Add(Column_unit_code)

                        Dim Column_Ent_dt As New Data.DataColumn("Ent_dt", GetType(System.DateTime))
                        Column_Ent_dt.DefaultValue = _currentDatetime
                        dtUploadExcel.Columns.Add(Column_Ent_dt)

                        Dim Column_Ent_Userid As New Data.DataColumn("Ent_Userid", GetType(System.String))
                        Column_Ent_Userid.DefaultValue = mP_User
                        dtUploadExcel.Columns.Add(Column_Ent_Userid)

                        Dim Column_Upd_dt As New Data.DataColumn("Upd_dt", GetType(System.DateTime))
                        Column_Upd_dt.DefaultValue = _currentDatetime
                        dtUploadExcel.Columns.Add(Column_Upd_dt)

                        Dim Column_Upd_Userid As New Data.DataColumn("Upd_Userid", GetType(System.String))
                        Column_Upd_Userid.DefaultValue = mP_User
                        dtUploadExcel.Columns.Add(Column_Upd_Userid)

                        Dim objbulk As SqlBulkCopy = New SqlBulkCopy(SqlConnectionclass.GetConnection())
                        objbulk.DestinationTableName = "EMPRO_DYN_" & txtTemplateUpload.Text & ""

                        'SqlConnectionclass.OpenConnection()
                        Using conn As SqlConnection = New SqlConnection("Server=" & gstrCONNECTIONSERVER & ";Database=" & gstrDatabaseName & ";user=" & gstrCONNECTIONUSER & ";password=" & gstrCONNECTIONPASSWORD & " ")
                            If conn.State = ConnectionState.Closed Then
                                conn.Open()
                            End If
                            Try
                                objbulk.BulkCopyTimeout = 180
                                objbulk.WriteToServer(dtUploadExcel)
                            Catch ex As Exception
                                If conn.State = ConnectionState.Open Then
                                    conn.Close()
                                End If
                                Throw ex
                            End Try
                        End Using
                        'SqlConnectionclass.CloseConnection()
                        SqlConnectionclass.ExecuteScalar("INSERT INTO GRN_TEMPLATE_UPLOAD_HISTORY(UPLOAD_ID,UPLOAD_TABLE_NAME,UPLOAD_DATE,UNIT_CODE,Ent_dt,Ent_Userid,Upd_dt,Upd_Userid,TEMPLATE_ID,Customer_Code,Upload_Type)VALUES('" & _currentDatetime.ToString("ddMMyyyyhhmmss") & "','EMPRO_DYN_" & txtTemplateUpload.Text & "',GETDATE(),'" & gstrUNITID & "',GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "',(Select TOP 1 TEMPLATE_ID from GRN_TEMPLATE_MST WHERE TEMPLATE_TABLE='EMPRO_DYN_" & txtTemplateUpload.Text & "'),'" & txtCustomerCodeUpload.Text & "','" & rbPaymentValue & "')")
                        MessageBox.Show("Data has been Imported successfully." & vbNewLine & " Upload ID is :- " & _currentDatetime.ToString("ddMMyyyyhhmmss") & "", "Imported", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        dtUploadExcel.Clear()
                        grdExcelData.DataSource = Nothing
                        txtTemplateUpload.Text = String.Empty
                        grdTempMatch.DataSource = Nothing
                        txtFileUpload.Text = String.Empty
                        txtCustomerCodeUpload.Text = String.Empty

                    End If
                End If
            Else
                MsgBox("Data is not available in grid to upload.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If
        Catch ex As Exception
            txtFileUpload.Text = String.Empty
            grdExcelData.DataSource = Nothing
            'SqlConnectionclass.CloseConnection()
            RaiseException(ex.InnerException)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnUploadClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUploadClose.Click
        Dim intResult As Integer = 0
        intResult = MsgBox("Do You Want To Close this Screen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100))
        If intResult = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub btnCustomerCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustomerCode.Click
        Dim strSql As String = String.Empty
        Dim strHelp() As String = Nothing
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            strSql = "select Customer_code,cust_name from Customer_mst WHERE UNIT_CODE='" & gstrUNITID & "'"
            strHelp = BtnHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, ResolveResString(100), 0)
            If UBound(strHelp) = -1 Then Exit Sub
            If strHelp(0) <> "0" Then
                txtCustomerCodeUpload.Text = strHelp(0).Trim
            Else
                MsgBox("No Record found !", MsgBoxStyle.Information, ResolveResString(100))
                txtTemplate_Sch.Text = String.Empty
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub cmdUploadID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUploadID.Click
        Dim strSql As String = String.Empty
        Dim strHelp() As String = Nothing
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            strSql = "SELECT UPLOAD_ID,UPLOAD_DATE FROM GRN_TEMPLATE_UPLOAD_HISTORY WHERE UNIT_CODE='" & gstrUNITID & "' ORDER BY UPLOAD_DATE"
            strHelp = BtnHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, ResolveResString(100), 0)
            If UBound(strHelp) = -1 Then Exit Sub
            If strHelp(0) <> "0" Then
                txtUploadID.Text = strHelp(0).Trim
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim _dtUploadedData As DataTable = SqlConnectionclass.GetDataTable("EXEC USP_GRIN_TEMPLATE_UPLOADED_DATA_GET '" & strHelp(0).Trim & "','" & gstrUNITID & "'")
                grdUploadedData.DataSource = _dtUploadedData
                grdUploadedData.AllowUserToAddRows = False
                grdUploadedData.Visible = True
            Else
                MsgBox("No Record found !", MsgBoxStyle.Information, ResolveResString(100))
                txtTemplate_Sch.Text = String.Empty
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnNewUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewUpload.Click
        btnCustomerCode.Enabled = True
        btnTemplateMatch.Enabled = True
        rbPayment.Enabled = True
        rbGRIN.Enabled = True
        btnFileUpload.Enabled = True
        grdTempMatch.Visible = True
        grdExcelData.Visible = True
        grdUploadedData.Visible = False
        grdUploadedData.DataSource = Nothing
        txtUploadID.Text = String.Empty
        txtUploadID.Enabled = False
        cmdUploadID.Enabled = False
        btnUploadExcel.Enabled = True
    End Sub
    Private Sub DisableUploadControl()
        Try
            txtCustomerCodeUpload.Enabled = False
            btnCustomerCode.Enabled = False
            txtTemplateUpload.Enabled = False
            btnTemplateMatch.Enabled = False
            rbPayment.Enabled = False
            rbGRIN.Enabled = False
            txtFileUpload.Enabled = False
            btnFileUpload.Enabled = False
            grdTempMatch.Visible = False
            grdExcelData.Visible = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdExceptionCustCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExceptionCustCode.Click
        Dim strSql As String = String.Empty
        Dim strHelp() As String = Nothing
        Dim _Upload_Type As String = Nothing
        Try
            If rbExceptionPayment.Checked = False AndAlso rbExceptionGRIN.Checked = False Then
                MsgBox("Upload type can not be uncheck.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim isChecked As Boolean = rbExceptionPayment.Checked
            If isChecked Then
                _Upload_Type = rbExceptionPayment.Text
            Else
                _Upload_Type = rbExceptionGRIN.Text
            End If

            strSql = "Select DISTINCT A.CUSTOMER_CODE Customer_code,B.cust_name from GRN_TEMPLATE_UPLOAD_HISTORY A INNER JOIN Customer_mst B ON A.CUSTOMER_CODE=B.Customer_Code where B.unit_code='" & gstrUNITID & "' AND Upload_Type='" & _Upload_Type & "'"
            strHelp = BtnHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, ResolveResString(100), 0)
            If UBound(strHelp) = -1 Then Exit Sub
            If strHelp(0) <> "0" Then
                txtExceptionCustCode.Text = strHelp(0).Trim
            Else
                MsgBox("No Record found !", MsgBoxStyle.Information, ResolveResString(100))
                txtExceptionCustCode.Text = String.Empty
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnExceptionClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExceptionClose.Click
        Dim intResult As Integer = 0
        intResult = MsgBox("Do You Want To Close this Screen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100))
        If intResult = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub btnExceptionDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExceptionDownload.Click
        Try
            Dim dtReturn As DataTable = New DataTable()
            Dim _Upload_Type As String = Nothing

            If rbExceptionPayment.Checked = False AndAlso rbExceptionGRIN.Checked = False Then
                MsgBox("Upload type can not be uncheck.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If
            If Len(txtExceptionCustCode.Text) = 0 Then
                MsgBox("Customer code can not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If

            Dim isChecked As Boolean = rbExceptionPayment.Checked
            If isChecked Then
                _Upload_Type = rbExceptionPayment.Text
            Else
                _Upload_Type = rbExceptionGRIN.Text
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "EXEC USP_ASN_GRIN_EXCEPTION_REPORT '" & txtExceptionCustCode.Text.Trim().Replace("'", "''") & "','" & gstrUNITID & "' ,'" & _Upload_Type & "'"
                    .CommandTimeout = 0
                    .CommandType = CommandType.Text
                    dtReturn = SqlConnectionclass.GetDataTable(sqlCmd)
                End With
            End Using

            'dtReturn = SqlConnectionclass.GetDataTable("EXEC USP_ASN_GRIN_EXCEPTION_REPORT '" & txtExceptionCustCode.Text.Trim().Replace("'", "''") & "','" & gstrUNITID & "' ,'" & _Upload_Type & "'")
            Dim savefile As SaveFileDialog = New SaveFileDialog()
            savefile.FileName = "Exception_Report_" & System.DateTime.Now.ToFileTime() & ".xls"
            savefile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            If dtReturn IsNot Nothing AndAlso dtReturn.Rows.Count > 0 Then

                If Convert.ToString(dtReturn.Rows(0).Item("MSG")) <> "SUCCESS" Then
                    Dim _retVal As String
                    For Each row As DataRow In dtReturn.Rows
                        _retVal += Convert.ToString(row("MSG")) + vbCrLf
                    Next
                    MsgBox(_retVal, MsgBoxStyle.Exclamation, ResolveResString(100))
                    Return
                End If
                dtReturn.Columns.Remove("MSG")
                If savefile.ShowDialog() = DialogResult.OK Then
                    Dim wr As StreamWriter = New StreamWriter(savefile.FileName)
                    For i As Integer = 0 To dtReturn.Columns.Count - 1
                        wr.Write(dtReturn.Columns(i).ToString().ToUpper() & vbTab)
                    Next
                    wr.WriteLine()
                    For i As Integer = 0 To (dtReturn.Rows.Count) - 1
                        For j As Integer = 0 To dtReturn.Columns.Count - 1
                            If dtReturn.Rows(i)(j) IsNot Nothing Then
                                wr.Write(Convert.ToString(dtReturn.Rows(i)(j)) & vbTab)
                            Else
                                wr.Write(vbTab)
                            End If
                        Next
                        wr.WriteLine()
                    Next
                    wr.Close()
                    MsgBox("Data has been downloaded in excel format at location " & savefile.FileName, MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            Else
                MsgBox("Can't export file because zero record to export.", MsgBoxStyle.OkOnly, ResolveResString(100))
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub rbExceptionPayment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbExceptionPayment.CheckedChanged
        txtExceptionCustCode.Text = String.Empty
    End Sub

    Private Sub rbExceptionGRIN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbExceptionGRIN.CheckedChanged
        txtExceptionCustCode.Text = String.Empty
    End Sub
End Class
