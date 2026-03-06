Imports System
Imports System.Data
Imports System.Data.SqlClient
'----------------------------------------------------
'Copyright(c)       - MIND
'Name of Module     - MARKETING CDR FUNCTIONALITY 
'Name of Form       - FRMMKTTRN0091.FRM  ,   INVOICE WISE ACTUAL RECEIPT AUTHORIZATION
'Created by         - Mayur Kumar
'Modified By        - 
'Created Date       - 16 SEP 2015
'description        - #10816097 -- New Forms Developed 
'*********************************************************************************************************************
Public Class FRMMKTTRN0091

    Private Enum ENUMINVOICEDETAILS
        VAL_APPROVE = 1
        VAL_REJECT
        VAL_INVOICENO
        VAL_ITEMCODE
        VAL_ITEMDESC
        VAL_INVOICEQTY
        VAL_SHORTQTY
        VAL_CUSTOMERQTY
        VAL_REMARKS
    End Enum

    Private Sub FRMMKTTRN0092_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, grpBoxGrid, ctlHeader, grpBoxButtons)
            Me.MdiParent = mdifrmMain

            cmdContract.Enabled = True
            btn_Save.Enabled = False

            chk_select.Enabled = False
            chk_remove.Enabled = False

            txt_ContractCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btn_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdContract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdContract.Click
        Try
            Dim strQuery As String
            Dim strHelp() As String

            strQuery = "SELECT DISTINCT DOC_NO,CONVERT(varchar(20),DOC_DATE,106) as DOC_DATE FROM SHORTRECEIPT_HDR WHERE UNIT_CODE='" + gstrUNITID + "' ORDER BY DOC_DATE DESC"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Document")
            If UBound(strHelp) > 0 Then
                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                    MsgBox("Document Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                If IsNothing(strHelp) = False Then
                    txt_ContractCode.Text = Convert.ToInt32(Trim(strHelp(0)))
                    txt_DocDate.Text = Trim(strHelp(1))
                    GETDOCUMENTDATA(Convert.ToInt32(Trim(strHelp(0))))
                    btn_Save.Enabled = True
                    'LockGrid()
                    chk_select.Enabled = True
                    chk_remove.Enabled = True
                End If

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub GETDOCUMENTDATA(ByVal doc_no As Integer)
        Dim strSql As String
        Dim dtb As New DataTable
        Dim con As SqlConnection
        Dim dad As SqlDataAdapter
        Try
            con = SqlConnectionclass.GetConnection()
            strSql = "SELECT INVOICE_NO,ITEM_CODE,ITEMDESC,INVOICEQTY,SHORTQTY,CUSTOMERQTY,REMARKS,STATUS FROM SHORTRECEIPT_DTL WHERE  UNIT_CODE='" + gstrUNITID + "' AND DOC_NO='" + doc_no.ToString() + "'"

            dad = New SqlDataAdapter(strSql, con)
            dad.Fill(dtb)
            InitializeSpread_InvoiceDtls()

            For Each row As DataRow In dtb.Rows

                AddRow()

                Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_INVOICENO
                Me.GridComp.Value = row("INVOICE_NO")
                Me.GridComp.Lock = True


                Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_ITEMCODE
                Me.GridComp.Value = row("ITEM_CODE")
                Me.GridComp.Lock = True

                Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_ITEMDESC
                Me.GridComp.Value = row("ITEMDESC")
                Me.GridComp.Lock = True

                Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_INVOICEQTY
                Me.GridComp.Value = row("INVOICEQTY")
                Me.GridComp.Lock = True

                Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_SHORTQTY
                Me.GridComp.Value = row("SHORTQTY")
                Me.GridComp.Lock = True

                Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_CUSTOMERQTY
                Me.GridComp.Value = row("CUSTOMERQTY")
                Me.GridComp.Lock = True

                Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_REMARKS
                Me.GridComp.Value = row("REMARKS")
                Me.GridComp.Lock = True

                If Convert.ToString(row("STATUS")).Equals("APPROVED") Then  ''if approved
                    Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_APPROVE
                    Me.GridComp.Text = 1
                    Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_REJECT
                    Me.GridComp.Text = 0
                    Me.GridComp.BlockMode = True
                    Me.GridComp.Row = Me.GridComp.MaxRows
                    Me.GridComp.Row2 = Me.GridComp.MaxRows
                    Me.GridComp.Col = 1
                    Me.GridComp.Col2 = Me.GridComp.MaxCols
                    Me.GridComp.Lock = True
                    Me.GridComp.BackColor = Color.GreenYellow
                    Me.GridComp.BlockMode = False
                ElseIf Convert.ToString(row("STATUS")).Equals("REJECTED") Then   ''If rejected
                    Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_APPROVE
                    Me.GridComp.Text = 0
                    Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_REJECT
                    Me.GridComp.Text = 1
                    Me.GridComp.BlockMode = True
                    Me.GridComp.Row = Me.GridComp.MaxRows
                    Me.GridComp.Row2 = Me.GridComp.MaxRows
                    Me.GridComp.Col = 1
                    Me.GridComp.Col2 = Me.GridComp.MaxCols
                    Me.GridComp.Lock = True
                    Me.GridComp.BackColor = Color.LightPink
                    Me.GridComp.BlockMode = False
                ElseIf Convert.ToString(row("STATUS")).Equals("SUBMITTED") Then  ''if nothing happened
                    Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_APPROVE
                    Me.GridComp.Text = 0
                    Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_REJECT
                    Me.GridComp.Text = 0
                    Me.GridComp.BlockMode = True
                    Me.GridComp.Row = Me.GridComp.MaxRows
                    Me.GridComp.Row2 = Me.GridComp.MaxRows
                    Me.GridComp.Col = ENUMINVOICEDETAILS.VAL_APPROVE
                    Me.GridComp.Col2 = ENUMINVOICEDETAILS.VAL_REJECT
                    Me.GridComp.Lock = False
                    Me.GridComp.BackColor = Color.White
                    Me.GridComp.BlockMode = False

                End If

            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub InitializeSpread_InvoiceDtls()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.GridComp
                .MaxRows = 0
                .MaxCols = ENUMINVOICEDETAILS.VAL_REMARKS
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_APPROVE : .Text = "APPROVE" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_REJECT : .Text = "REJECT" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_INVOICENO : .Text = "INVOICE NO" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_ITEMCODE : .Text = "ITEM CODE" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_ITEMDESC : .Text = "ITEM DESCRIPTION" : .set_ColWidth(.Col, 20) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_INVOICEQTY : .Text = "INVOICE QTY" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_SHORTQTY : .Text = "SHORT RECEIPT" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_CUSTOMERQTY : .Text = "CUSTOMER RECEIPT" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_REMARKS : .Text = "REMARKS" : .set_ColWidth(.Col, 25) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .BlockMode = True
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Row2 = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = 1
                .Col2 = .MaxCols
                .Lock = True
                .BlockMode = False
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Public Sub AddRow()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

            With GridComp
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_APPROVE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_REJECT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_INVOICENO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_ITEMCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_ITEMDESC : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_INVOICEQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_CUSTOMERQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_SHORTQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_REMARKS : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)

            End With

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub GridComp_ClickEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles GridComp.ClickEvent
        Try
            If GridComp.MaxRows > 0 Then
                If e.col = ENUMINVOICEDETAILS.VAL_APPROVE Then
                    With GridComp
                        .Row = e.row
                        .Col = ENUMINVOICEDETAILS.VAL_REJECT
                        If .Lock = False Then
                            .Value = 0
                        End If
                    End With
                ElseIf e.col = ENUMINVOICEDETAILS.VAL_REJECT Then
                    With GridComp
                        .Row = e.row
                        .Col = ENUMINVOICEDETAILS.VAL_APPROVE
                        If .Lock = False Then
                            .Value = 0
                        End If
                    End With
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub GridComp_KeyPressEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles GridComp.KeyPressEvent
        Try
            If GridComp.ActiveCol = ENUMINVOICEDETAILS.VAL_APPROVE Or GridComp.ActiveCol = ENUMINVOICEDETAILS.VAL_REJECT Then
                If e.keyAscii = Keys.Space Then
                    GridComp_ClickEvent(GridComp, New AxFPSpreadADO._DSpreadEvents_ClickEvent(GridComp.ActiveCol, GridComp.ActiveRow))
                End If
            ElseIf e.keyAscii = 34 Or e.keyAscii = 39 Then
                e.keyAscii = 0
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btn_Reject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click
        Try
            Dim IntX As Integer
            Dim IntY As Integer
            Dim IsTrans As Boolean = False
            Dim isRecordSaved As Boolean = False
            Dim Sqlcmd As New SqlCommand
            Dim SqlTrans As SqlTransaction
            Dim strSql As String = String.Empty
            Dim IsAuthorized As Boolean
            Dim status As String = String.Empty
            Dim Invoice As String = String.Empty
            Dim ItemCode As String = String.Empty

            Sqlcmd.Connection = SqlConnectionclass.GetConnection()
            If (GridComp.MaxRows <= 0) Then
                MessageBox.Show("Please Select Document.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            SqlTrans = Sqlcmd.Connection.BeginTransaction
            Sqlcmd.Transaction = SqlTrans
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.CommandType = CommandType.Text
            IsTrans = True

            strSql = "UPDATE SHORTRECEIPT_DTL SET STATUS =@STATUS,UPD_DT=GETDATE(),UPD_USERID=@USERID WHERE DOC_NO =@DOCNO AND UNIT_CODE =@UNIT_CODE AND INVOICE_NO =@INVOICE_NO AND ITEM_CODE =@ITEM_CODE"

            For IntX = 1 To Me.GridComp.MaxRows
                GridComp.Row = IntX
                GridComp.Col = ENUMINVOICEDETAILS.VAL_APPROVE
                If GridComp.Lock = True Then
                    Continue For
                End If
                If GridComp.Value = 0 Then
                    GridComp.Col = ENUMINVOICEDETAILS.VAL_REJECT
                    IsAuthorized = GridComp.Value
                    If GridComp.Value = 0 Then
                        Continue For
                    Else
                        IsAuthorized = False
                    End If
                Else
                    IsAuthorized = True
                End If


                GridComp.Col = ENUMINVOICEDETAILS.VAL_INVOICENO
                Invoice = GridComp.Value.Trim

                GridComp.Col = ENUMINVOICEDETAILS.VAL_ITEMCODE
                ItemCode = GridComp.Value.Trim


                Sqlcmd.CommandText = strSql
                Sqlcmd.Parameters.Clear()
                Sqlcmd.Parameters.AddWithValue("@DOCNO", txt_ContractCode.Text.Trim())
                Sqlcmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                Sqlcmd.Parameters.AddWithValue("@ITEM_CODE", ItemCode)
                Sqlcmd.Parameters.AddWithValue("@INVOICE_NO", Invoice)

                If IsAuthorized = False Then
                    status = "REJECTED"
                End If
                If IsAuthorized = True Then
                    status = "APPROVED"
                End If

                Sqlcmd.Parameters.AddWithValue("@STATUS", status)
                Sqlcmd.Parameters.AddWithValue("@USERID", mP_User)
                Sqlcmd.ExecuteNonQuery()
                isRecordSaved = True

            Next
            If (isRecordSaved) Then
                If MessageBox.Show("Are you sure, you want to save?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    SqlTrans.Rollback()
                    IsTrans = False
                    Exit Sub
                End If
                SqlTrans.Commit()
                IsTrans = False
                MessageBox.Show("Record save Successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                GETDOCUMENTDATA(Convert.ToInt32(txt_ContractCode.Text.Trim()))
            Else
                MessageBox.Show("Nothing Selected.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub chk_select_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_select.CheckedChanged
        Try
            Dim intloopcounter As Int16 = 0

            If chk_remove.Checked = True Then
                'chk_remove.Checked = False
            End If

            If chk_select.Checked = True Then
                With GridComp
                    .Col = ENUMINVOICEDETAILS.VAL_APPROVE
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 1
                        End If
                    Next
                    .Col = ENUMINVOICEDETAILS.VAL_REJECT
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 0
                        End If
                    Next
                End With
            End If

            If chk_select.Checked = False Then
                With GridComp
                    .Col = ENUMINVOICEDETAILS.VAL_APPROVE
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 0
                        End If
                    Next
                End With
            End If

            'chk_remove.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub chk_remove_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_remove.CheckedChanged
        Try
            Dim intloopcounter As Int16 = 0

            If chk_select.Checked = True Then
                'chk_select.Checked = False
            End If

            If chk_remove.Checked = True Then
                With GridComp
                    .Col = ENUMINVOICEDETAILS.VAL_REJECT
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 1
                        End If
                    Next
                    .Col = ENUMINVOICEDETAILS.VAL_APPROVE
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 0
                        End If
                    Next
                End With
            End If

            If chk_remove.Checked = False Then
                With GridComp
                    .Col = ENUMINVOICEDETAILS.VAL_REJECT
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 0
                        End If
                    Next
                End With
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class