Imports System
Imports System.Globalization
Imports System.Data.SqlClient

'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - ASN Generation Via API
'Name of Form       - FRMMKTTRN0083A  , ASN Generation Via API
'Created by         - Santosh Kumar Yadav
'Created Date       - 30 Sept 2022
'description        - ASN Generation Via API (New Development)
'*********************************************************************************************************************


Public Class FRMMKTTRN0083A
    Private Const Invoice As String = "INVOICES"
    Private Enum enumInvoiceDetail
        ProcessID = 1
        SLNO
        InvocieNo
        InvoiceDate
        CustomerCode
        CustomerName
        APIStatus
        APIMessage
        ReceivedASNNo
        Unit_Code
        IsProcess
    End Enum
    Private Sub FRMMKTTRN0083A_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        dtpDateFrom.Focus()
    End Sub
    Private Sub FRMMKTTRN0083A_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ' Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 600)
            dtpDateFrom.Value = GetServerDate()
            dtpDateFrom.Value = dtpDateFrom.Value.AddDays(-5)
            dtpDateTo.Value = GetServerDate()
            InitializeSpread()
            FillInvoices()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            End If

            FillInvoices()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub


    Private Sub dtpDateFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDateFrom.ValueChanged
        Try

            fspASNInvocieDetails.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dtpDateTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDateTo.ValueChanged
        Try
            fspASNInvocieDetails.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub



    Private Sub FillInvoices()
        Dim dt As New DataTable
        Try
            Dim sqlCmd As New SqlCommand
            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 300 ' 5 Minute
                .CommandText = "USP_TML_TCS_ASN_WebAPI_IntegrationDataGet"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UnitCode", gstrUNITID)
                .Parameters.AddWithValue("@FrmInvoiceDate", dtpDateFrom.Value)
                .Parameters.AddWithValue("@ToInvoiceDate", dtpDateTo.Value)
                dt = SqlConnectionclass.GetDataTable(sqlCmd)
                InitializeSpread()
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

                    For i As Integer = 0 To dt.Rows.Count - 1
                        With fspASNInvocieDetails
                            AddRow()
                            .SetText(enumInvoiceDetail.ProcessID, i + 1, dt.Rows(i).Item("ProcessID").ToString.Trim)
                            .SetText(enumInvoiceDetail.SLNO, i + 1, dt.Rows(i).Item("SNo"))
                            .SetText(enumInvoiceDetail.InvocieNo, i + 1, dt.Rows(i).Item("InvoiceNo"))
                            .SetText(enumInvoiceDetail.InvoiceDate, i + 1, dt.Rows(i).Item("InvoiceDate"))
                            .SetText(enumInvoiceDetail.CustomerCode, i + 1, dt.Rows(i).Item("Customer_Code"))
                            .SetText(enumInvoiceDetail.CustomerName, i + 1, dt.Rows(i).Item("CustomerName"))
                            .SetText(enumInvoiceDetail.APIStatus, i + 1, dt.Rows(i).Item("APIStatus"))
                            .SetText(enumInvoiceDetail.APIMessage, i + 1, dt.Rows(i).Item("APIMessage"))
                            .SetText(enumInvoiceDetail.ReceivedASNNo, i + 1, dt.Rows(i).Item("ReceivedASNNo"))
                            .SetText(enumInvoiceDetail.Unit_Code, i + 1, dt.Rows(i).Item("Unit_Code"))
                            .SetText(enumInvoiceDetail.IsProcess, i + 1, dt.Rows(i).Item("IsProcessed"))
                        End With
                    Next
                End If

            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dt.Dispose()
        End Try
    End Sub

    Private Sub InitializeSpread()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.fspASNInvocieDetails
                .MaxRows = 0
                .MaxCols = [Enum].GetValues(GetType(enumInvoiceDetail)).Length
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = enumInvoiceDetail.ProcessID : .Text = "Process Id" : .set_ColWidth(.Col, 18) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.SLNO : .Text = "SNo." : .set_ColWidth(.Col, 5) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.InvocieNo : .Text = "Invoice No." : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.InvoiceDate : .Text = "Invoice Date" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.CustomerCode : .Text = "Customer Code" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.CustomerName : .Text = "Customer Name" : .set_ColWidth(.Col, 20) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.APIStatus : .Text = "API Staus" : .set_ColWidth(.Col, 6) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.APIMessage : .Text = "API Message" : .set_ColWidth(.Col, 25) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.ReceivedASNNo : .Text = "Received ASN No" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.Unit_Code : .Text = "Unit Code" : .set_ColWidth(.Col, 5) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumInvoiceDetail.IsProcess : .Text = "Processed Status" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
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
            With fspASNInvocieDetails
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = enumInvoiceDetail.ProcessID : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.SLNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.InvocieNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.InvoiceDate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY
                .Row = .MaxRows : .Col = enumInvoiceDetail.CustomerCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.CustomerName : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.APIStatus : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .FontBold = True : .BackColor = Color.Gray
                .Row = .MaxRows : .Col = enumInvoiceDetail.APIMessage : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.ReceivedASNNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.Unit_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumInvoiceDetail.IsProcess : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class