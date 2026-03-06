Imports System
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

    ''Revised By:       Saurav Kumar
    ''Revised On:       04 Oct 2013
    ''Issue ID  :       10462231 - eMpro ISuite Changes
'***********************************************************************************************************************************
''Changed By:       Mayur Kumar
''Changed On:       09 July 2015
''Issue ID  :       10854727  - Grid Changed due to no of rows fixed in earlier vb grid.  Full Form Changed.
'***********************************************************************************************************************************

Friend Class FRMMKTTRN0073A
    Inherits System.Windows.Forms.Form
    Dim dtPRDtl As DataTable

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Try
            Me.Close()
            Me.Dispose()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Try
            If dgvPRDtl.Rows.Count > 0 Then
                Call FlexGrid_To_Excel("Export")
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub Form3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            SetBackGroundColorNew(Me, True)
            cmblist1.Items.Clear()
            cmblist1.Items.Add(("UB"))
            cmblist1.Items.Add(("MK"))
            cmblist1.Items.Add(("DL"))
            AddColumnPRDtlGrid()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub AddColumnPRDtlGrid()
        Try
            If dgvPRDtl.Columns.Count = 0 Then

                Dim dgvA As New DataGridViewTextBoxColumn
                dgvA.DataPropertyName = "TRIG_LOC"
                dgvA.Name = "TRIG_LOC"
                dgvA.HeaderText = "Trigger Location"
                dgvA.Width = 60
                dgvA.ReadOnly = True
                dgvPRDtl.Columns.Add(dgvA)

                Dim dgvC As New DataGridViewTextBoxColumn
                dgvC.DataPropertyName = "CUST_CODE"
                dgvC.Name = "CUST_CODE"
                dgvC.HeaderText = "Customer Code"
                dgvC.Width = 110
                dgvC.ReadOnly = True
                dgvPRDtl.Columns.Add(dgvC)

                Dim dgvD As New DataGridViewTextBoxColumn
                dgvD.DataPropertyName = "CAT_CODE"
                dgvD.Name = "CAT_CODE"
                dgvD.HeaderText = "Cat Code"
                dgvD.Width = 100
                dgvD.ReadOnly = True
                dgvPRDtl.Columns.Add(dgvD)

                Dim dgvQ As New DataGridViewTextBoxColumn
                dgvQ.DataPropertyName = "BODY_COLOR"
                dgvQ.Name = "BODY_COLOR"
                dgvQ.HeaderText = "Body Color"
                dgvQ.Width = 80
                dgvQ.ReadOnly = True
                dgvPRDtl.Columns.Add(dgvQ)

                Dim dgvN As New DataGridViewTextBoxColumn
                dgvN.DataPropertyName = "MODEL_NO"
                dgvN.Name = "MODEL_NO"
                dgvN.HeaderText = "Model NO"
                dgvN.Width = 110
                dgvN.ReadOnly = True
                dgvPRDtl.Columns.Add(dgvN)

                Dim dgvUM As New DataGridViewTextBoxColumn
                dgvUM.DataPropertyName = "QTY"
                dgvUM.Name = "QTY"
                dgvUM.HeaderText = "QTY"
                dgvUM.Width = 47
                dgvUM.ReadOnly = True
                dgvPRDtl.Columns.Add(dgvUM)


                dgvPRDtl.AutoGenerateColumns = False
                dgvPRDtl.RowsDefaultCellStyle.BackColor = Color.Lavender
                dgvPRDtl.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(190, 200, 255)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
  
    Private Sub ClearFlex()
        Try
            If dgvPRDtl.Rows.Count > 0 Then
                dgvPRDtl.DataSource = Nothing
                dgvPRDtl.Refresh()
                dgvPRDtl.Rows.Clear()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub REFRESH_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles REFRESH_Renamed.Click
        Try
            Call ClearFlex()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub view_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles view.Click
        Try
            If Trim(Me.cmblist1.Text) = "" Then
                MsgBox("You need to Select the Trigger Location", MsgBoxStyle.Critical)
                Exit Sub
            Else
                fillPRItem()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Function fillPRItem() As Boolean
        Dim sqlQuery As String = String.Empty
        Dim Sqlcmd As New SqlCommand
        Dim SQLAdp As SqlDataAdapter
        Try

            If String.IsNullOrEmpty(sqlQuery) Then
                sqlQuery = "SELECT * FROM GETTRIGGERDTL('" & gstrUNITID & "','" & Trim(Me.cmblist1.Text) & "','" & Trim(getDateForDB(Me.DTPicker1.Value)) & "','" & Trim(getDateForDB(Me.DTPicker2.Value)) & "')"
            End If
            dtPRDtl = New DataTable
            With Sqlcmd
                .CommandType = CommandType.Text
                .CommandText = sqlQuery
                .Connection = SqlConnectionclass.GetConnection()
            End With
            SQLAdp = New SqlDataAdapter(Sqlcmd)
            SQLAdp.Fill(dtPRDtl)
            dgvPRDtl.DataSource = dtPRDtl
            For col As Integer = 1 To dgvPRDtl.Columns.Count - 1
                dgvPRDtl.Columns(col).ReadOnly = True
            Next

        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return True
    End Function

    Public Sub FlexGrid_To_Excel(Optional ByRef WorkSheetName As String = "")
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

            Dim objXL As New Excel.Application
            Dim wbXL As Excel.Workbook
            Dim wsXL As New Excel.Worksheet



            Dim intRow As Short ' counter
            Dim intCol As Short ' counter

            If Not IsReference(objXL) Then
                MsgBox("You need Microsoft Excel to use this function", MsgBoxStyle.Exclamation, "Print to Excel")
                Exit Sub
            End If

            objXL.Visible = False
            wbXL = objXL.Workbooks.Add
            wsXL = objXL.ActiveSheet

            With wsXL
                If Not WorkSheetName = "" Then
                    .Name = WorkSheetName
                End If
            End With


            Dim column As DataGridViewColumn
            For Each column In dgvPRDtl.Columns
                wsXL.Cells(1, column.Index + 1).Value = column.HeaderText
            Next

            For i As Integer = 0 To dgvPRDtl.Rows.Count - 1
                Dim columnIndex As Integer = 0
                Do Until columnIndex = dgvPRDtl.Columns.Count
                    wsXL.Cells(i + 2, columnIndex + 1).Value = dgvPRDtl.Item(columnIndex, i).Value.ToString
                    columnIndex += 1
                Loop
            Next

            objXL.Visible = True

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class