'--------------------------------------------------------------------------------------------------
'COPYRIGHT      :   MIND
'CREATED BY     :   SUMIT KUMAR
'CREATED DATE   :   25 SEPT 2019
'SCREEN         :   BSR Free Bag Label

'--------------------------------------------------------------------------------------------------

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Collections.Generic

Public Class frmMKTTRN0111

    Dim mintFormIndex As Integer


    Private Enum enmPickList
        Item_Code = 1
        Barcode = 2
        InvoiceNO = 3
        CloseBox = 4
        Qty = 5
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
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
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
            Call FitToClient(Me, GrpMain, ctlHeader, grpBtn, 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            SetGridsHeader()
            txtscanbarcode_view.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Methods"
    Private Sub SetGridsHeader()
        Try

            With sprPicklist
                .MaxRows = 0
                .MaxCols = [Enum].GetNames(GetType(enmPickList)).Count
                .Row = 0
                .set_RowHeight(0, 20)
                .Col = 0 : .set_ColWidth(0, 3)
                .Col = enmPickList.Item_Code : .Text = "Item Code" : .set_ColWidth(enmPickList.Item_Code, 16)
                .Col = enmPickList.Barcode : .Text = "Bar Code" : .set_ColWidth(enmPickList.Barcode, 40)
                .Col = enmPickList.InvoiceNO : .Text = "Invoice No" : .set_ColWidth(enmPickList.InvoiceNO, 10)
                .Col = enmPickList.CloseBox : .Text = "Close BoxNo" : .set_ColWidth(enmPickList.CloseBox, 8)
                .Col = enmPickList.Qty : .Text = "Qty" : .set_ColWidth(enmPickList.Qty, 8)
              
            End With


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddBlankRow()
        Try
            With Me.sprPicklist
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
                .set_RowHeight(.Row, 15)
                .Col = enmPickList.Item_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

                .Col = enmPickList.Barcode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enmPickList.InvoiceNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enmPickList.CloseBox : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enmPickList.Qty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
   
#End Region


   

   

  

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Try
            txtscanbarcode_view.Text = ""
            txtFreeBarcode.Text = ""
            txtboxno.Text = ""
            txtInvoiceNo.Text = ""
            lblviewbarcode.Text = ""
            lblBarcode_freelabel.Text = ""
            txtbarqty.Text = ""
            sprPicklist.MaxRows = 0
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

   
    Private Sub btnMakeFree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMakeFree.Click
        Try
           
            If lblBarcode_freelabel.Text.Trim = "" Then
                MsgBox("Please Scan Valid Barcode!", MsgBoxStyle.Information, ResolveResString(100))
                txtFreeBarcode.Text = String.Empty
                txtFreeBarcode.Focus()
                Exit Sub
            End If
            If MessageBox.Show("Are you Sure To Make Free?", "Confirmation", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            If txtbarqty.Text.Trim = "" Or txtbarqty.Text.Trim = "0" Then
                Dim arrbarcode As String()
                arrbarcode = lblBarcode_freelabel.Text.Split("|")
                txtbarqty.Text = arrbarcode(1).ToString()
            End If

            Dim Qselect As String = String.Empty
            Qselect = "EXEC USP_BSR_VIEW_FREE_BAG_LABEL_BYBARCODE '" & gstrUNITID.ToString() & "', '" & lblBarcode_freelabel.Text.Trim.ToString() & "','SAVE_FREE_SCANNED_BAR_CODE','" & txtboxno.Text.Trim.ToString() & "','','" & txtbarqty.Text.Trim.ToString() & "','" & mP_User.ToString() & "'"

            Dim da As SqlDataAdapter = New SqlDataAdapter(Qselect, SqlConnectionclass.GetConnection)
            Dim dt As New DataTable
            da.SelectCommand.CommandTimeout = 120
            da.Fill(dt)

            If dt.Rows.Count = 0 Then
                MsgBox("ERROR TO MAKE FREE BARCODE!", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            Else
               MessageBox.Show(Convert.ToString(dt.Rows(0).Item(0)), "EMPRO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtFreeBarcode.Text = ""
                txtFreeBarcode.Focus()
                lblBarcode_freelabel.Text = ""
                txtboxno.Text = ""
                txtInvoiceNo.Text = ""
                txtbarqty.Text = ""

            End If
            txtFreeBarcode.Text = ""


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Dim dtGlbl As DataTable
    Private Sub txtscanbarcode_view_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtscanbarcode_view.KeyDown
        lblviewbarcode.Text = ""
        sprPicklist.MaxRows = 0
        Dim oCMD As SqlCommand
        Dim Err_msg As String=""
        Try
            If e.KeyValue = 13 Then
                If txtscanbarcode_view.Text.Trim() = "" Then
                    MsgBox("Please Scan Valid Barcode!", MsgBoxStyle.Information, ResolveResString(100))
                    txtscanbarcode_view.Text = ""
                    txtscanbarcode_view.Focus()
                    Exit Sub
                End If

                Dim Qselect As String = String.Empty
                Qselect = "EXEC USP_BSR_VIEW_FREE_BAG_LABEL_BYBARCODE '" & gstrUNITID.ToString() & "', '" & txtscanbarcode_view.Text.Trim.ToString() & "','VIEW_SCANNED_BAR_CODE','','','','" & mP_User.ToString() & "'"

                Dim da As SqlDataAdapter = New SqlDataAdapter(Qselect, SqlConnectionclass.GetConnection)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 120
                da.Fill(dt)

                If dt.Rows.Count = 0 Or dt.Columns.Count = 1 Then
                    MessageBox.Show(Convert.ToString(dt.Rows(0).Item(0)), "EMPRO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtscanbarcode_view.Text = ""
                    txtscanbarcode_view.Focus()
                    Exit Sub
                Else
                    lblviewbarcode.Text = txtscanbarcode_view.Text.Trim.ToString()
                    dtGlbl = dt.Copy
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With sprPicklist
                            .MaxRows = i + 1

                            .Row = .MaxRows : .Col = enmPickList.Item_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dt.Rows(i)("Item_Code").ToString <> "", dt.Rows(i)("Item_Code").ToString, String.Empty)
                            .Row = .MaxRows : .Col = enmPickList.Barcode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dt.Rows(i)("BARCODE").ToString <> "", dt.Rows(i)("BARCODE").ToString, String.Empty)
                            .Row = .MaxRows : .Col = enmPickList.InvoiceNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dt.Rows(i)("INVOICENO").ToString <> "", dt.Rows(i)("INVOICENO").ToString, String.Empty)
                            .Row = .MaxRows : .Col = enmPickList.CloseBox : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dt.Rows(i)("BOXNO").ToString <> "", dt.Rows(i)("BOXNO").ToString, String.Empty)
                            .Row = .MaxRows : .Col = enmPickList.Qty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dt.Rows(i)("QTY").ToString <> "", dt.Rows(i)("QTY").ToString, String.Empty)

                            .set_RowHeight(.MaxRows, 15)

                        End With
                    Next
                    txtFreeBarcode.Text = ""
                    txtFreeBarcode.Focus()

                End If
                txtscanbarcode_view.Text = ""


               
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

           
    End Sub

    Private Sub txtFreeBarcode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFreeBarcode.KeyDown
        lblBarcode_freelabel.Text = ""
        txtboxno.Text = ""
        txtInvoiceNo.Text = ""
        txtbarqty.Text = ""
        Dim oCMD As SqlCommand
        Dim Err_msg As String = ""
        Try
            If e.KeyValue = 13 Then
                If txtFreeBarcode.Text.Trim() = "" Then
                    MsgBox("Please Scan Valid Barcode!", MsgBoxStyle.Information, ResolveResString(100))
                    txtFreeBarcode.Text = ""
                    txtFreeBarcode.Focus()
                    Exit Sub
                End If

                Dim Qselect As String = String.Empty
                Qselect = "EXEC USP_BSR_VIEW_FREE_BAG_LABEL_BYBARCODE '" & gstrUNITID.ToString() & "', '" & txtFreeBarcode.Text.Trim.ToString() & "','GET_FREE_SCANNED_BAR_CODE','','','','" & mP_User.ToString() & "'"

                Dim da As SqlDataAdapter = New SqlDataAdapter(Qselect, SqlConnectionclass.GetConnection)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 120
                da.Fill(dt)

                If dt.Rows.Count = 0 Or dt.Columns.Count = 1 Then
                    MessageBox.Show(Convert.ToString(dt.Rows(0).Item(0)), "EMPRO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtFreeBarcode.Text = ""
                    txtFreeBarcode.Focus()
                    Exit Sub
                Else
                    lblBarcode_freelabel.Text = txtFreeBarcode.Text.Trim.ToString()
                    dtGlbl = dt.Copy
                    txtInvoiceNo.Text = Convert.ToString(dt.Rows(0).Item("INVOICENO"))
                    txtboxno.Text = Convert.ToString(dt.Rows(0).Item("BoxNo"))
                    txtbarqty.Text = Convert.ToString(dt.Rows(0).Item("Qty"))
                    
                    txtFreeBarcode.Text = ""
                    txtFreeBarcode.Focus()

                End If
                txtFreeBarcode.Text = ""



            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class





