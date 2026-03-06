Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
'----------------------------------------------------
'Copyright(c)       - MIND
'Name of Module     - SALES AND MARKETING
'Name of Form       - FRMMKTREP0071.FRM  ,   
'Created by         - Mayur Kumar
'Modified By        - 
'Created Date       - 16 Feb 2017
'description        - New Development - Transfer Invoice Label Printing
'*********************************************************************************************************************
Public Class FRMMKTTRN0093

#Region "GLOBAL VARIABLES"

    Dim SqlAdp As SqlDataAdapter
    Dim DSLBLDTL As DataSet
    Dim intLoopCounter As Int16 = 0
    Dim mstrErrMsg As String = String.Empty

    Private Enum ENUMLABELDETAILS
        VAL_SELECT = 1
        VAL_ITEMCODE
        VAL_ITEMDESC
        VAL_CUSTITEMCODE
        VAL_LABELS
        VAL_PACKSIZE
    End Enum

    Dim VAL_SELECT As Object = Nothing, VAL_LABELS As Object = Nothing
    Dim VAL_CUSTITEMCODE As Object = Nothing, VAL_ITEMDESC As Object = Nothing
    Dim VAL_ITEMCODE As Object = Nothing, VAL_PACKSIZE As Object = Nothing

#End Region

#Region "FORM EVENTS"

    Private Sub FRMMKTREP0071_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, fraContainer, ctlHeader, grpcmdbtns, 500)
            Me.MdiParent = mdifrmMain

            txtboxUnit.Focus()

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub
    Protected Overrides Sub WndProc(ByRef message As Message)

        Const WM_SYSCOMMAND As Integer = &H112
        Const SC_MOVE As Integer = &HF010

        Select Case message.Msg
            Case WM_SYSCOMMAND
                Dim command As Integer = message.WParam.ToInt32() And &HFFF0
                If command = SC_MOVE Then
                    Return
                End If
                Exit Select
        End Select

        MyBase.WndProc(message)

    End Sub

#End Region

#Region "BUTTON EVENTS"
    Private Sub btn_UNit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_UNit.Click
        Try
            Dim strQuery As String = String.Empty
            Dim strHelp() As String

            strQuery = "SELECT DISTINCT UNITCODE,VENDORCODE FROM ASN_LEBELS_TEMP WHERE VENDORCODE='" + gstrUNITID + "' "
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "UNIT")

            If UBound(strHelp) > 0 Then
                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                    MsgBox("Unit Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                    txtboxUnit.Focus()
                    Exit Sub
                End If
                If IsNothing(strHelp) = False Then
                    Me.txtboxUnit.Text = Trim(strHelp(0))
                    btnInvoiceNo.Enabled = True
                    txt_InvoiceNo.Focus()
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btnInvoiceNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvoiceNo.Click
        Try
            If (txtboxUnit.Text.Trim() <> "") Then

                Dim strQuery As String = String.Empty
                Dim strHelp() As String

                strQuery = "SELECT DISTINCT INVOICENO,INVOICEDATE FROM ASN_LEBELS_TEMP WHERE VENDORCODE='" + gstrUNITID + "' AND UNITCODE='" + txtboxUnit.Text.Trim.ToString() + "' "
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Invoice")
                If UBound(strHelp) > 0 Then
                    If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                        MsgBox("Invoice Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                        txtboxUnit.Focus()
                        Exit Sub
                    End If
                    If IsNothing(strHelp) = False Then
                        Me.txt_InvoiceNo.Text = Trim(strHelp(0))
                        If (txt_InvoiceNo.Text.Trim() <> "") Then
                            GETINVOICEDATA(txtboxUnit.Text.Trim(), txt_InvoiceNo.Text.Trim())
                            chk_Select.Checked = True
                        End If
                    End If
                End If

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtboxUnit_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtboxUnit.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                btn_UNit_Click(sender, e)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txt_InvoiceNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_InvoiceNo.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                btnInvoiceNo_Click(sender, e)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region

#Region "GRID EVENTS"

    Private Sub InitializeSpread_VendorDtls()

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.fspgrid_Contractdtls
                .MaxRows = 0
                .MaxCols = ENUMLABELDETAILS.VAL_PACKSIZE
                .set_RowHeight(0, 20)

                .Row = 0 : .Col = ENUMLABELDETAILS.VAL_SELECT : .Text = "SELECT" : .set_ColWidth(.Col, 4) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMLABELDETAILS.VAL_ITEMCODE : .Text = "ITEM CODE" : .set_ColWidth(.Col, 15) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMLABELDETAILS.VAL_ITEMDESC : .Text = "DESCRIPTION" : .set_ColWidth(.Col, 45) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMLABELDETAILS.VAL_CUSTITEMCODE : .Text = "CUSTOMER ITEM CODE" : .set_ColWidth(.Col, 15) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMLABELDETAILS.VAL_LABELS : .Text = "LABELS" : .set_ColWidth(.Col, 14) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMLABELDETAILS.VAL_PACKSIZE : .Text = "PACK SIZE" : .set_ColWidth(.Col, 5) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)

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
            With fspgrid_Contractdtls
                .MaxRows = .MaxRows + 1

                .Row = .MaxRows : .Col = ENUMLABELDETAILS.VAL_SELECT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMLABELDETAILS.VAL_ITEMCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMLABELDETAILS.VAL_ITEMDESC : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMLABELDETAILS.VAL_CUSTITEMCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMLABELDETAILS.VAL_LABELS : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMLABELDETAILS.VAL_PACKSIZE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)

            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

#End Region

#Region "DATA EVENTS"

    Private Sub GETINVOICEDATA(ByRef customer_code As String, ByRef invoiceno As String)

        Dim sqlCmd As New SqlCommand()
        SqlAdp = New SqlDataAdapter
        DSLBLDTL = New DataSet

        Try

            With sqlCmd
                .CommandText = "USP_LABELSDATA"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "VIEW")
                .Parameters.AddWithValue("@CUSTOMERCODE", customer_code.Trim.ToString())
                .Parameters.AddWithValue("@INVOICENO", invoiceno)
                .Parameters.AddWithValue("@VEDNOR_CODE", gstrUNITID)
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSLBLDTL)
                .Dispose()
            End With

            If DSLBLDTL.Tables.Count > 0 Then
                If DSLBLDTL.Tables(0).Rows.Count > 0 Then
                    InitializeSpread_VendorDtls()
                    With Me.fspgrid_Contractdtls
                        For intLoopCounter = 0 To DSLBLDTL.Tables(0).Rows.Count - 1
                            AddRow()
                            .SetText(ENUMLABELDETAILS.VAL_SELECT, intLoopCounter + 1, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("SELECT").ToString.Trim)
                            .SetText(ENUMLABELDETAILS.VAL_ITEMCODE, intLoopCounter + 1, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("Item_Code").ToString.Trim)
                            .SetText(ENUMLABELDETAILS.VAL_ITEMDESC, intLoopCounter + 1, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("Description"))
                            .SetText(ENUMLABELDETAILS.VAL_CUSTITEMCODE, intLoopCounter + 1, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("ItemCode"))
                            .SetText(ENUMLABELDETAILS.VAL_LABELS, intLoopCounter + 1, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("Label"))
                            .SetText(ENUMLABELDETAILS.VAL_PACKSIZE, intLoopCounter + 1, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("PackSize"))
                        Next
                    End With
                End If

            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
        End Try

    End Sub

#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try

            If txtboxUnit.Text = "" Then
                MsgBox("Kindly select Unit First!")
                Exit Sub
            End If

            If txt_InvoiceNo.Text = "" Then
                MsgBox("Kindly select Invoice First.")
                Exit Sub
            End If

            Dim label As String = String.Empty
            Dim count As Int16 = 0

            With Me.fspgrid_Contractdtls

                For intLoopCounter = 1 To .MaxRows

                    VAL_SELECT = Nothing
                    .GetText(ENUMLABELDETAILS.VAL_SELECT, intLoopCounter, VAL_SELECT)
                    If IsNothing(VAL_SELECT) = True Then VAL_SELECT = String.Empty

                    If VAL_SELECT = "1" Then

                        VAL_LABELS = Nothing
                        .GetText(ENUMLABELDETAILS.VAL_LABELS, intLoopCounter, VAL_LABELS)
                        If IsNothing(VAL_LABELS) = True Then VAL_LABELS = String.Empty

                        label = label + VAL_LABELS + ";"

                        count = count + 1
                    End If
                Next
            End With

            If count = 0 Then
                MsgBox("Kindly select the labels to print!")
                Exit Sub
            End If

            If count > 0 Then

                Dim sqlCmd As New SqlCommand()

                With sqlCmd
                    .CommandText = "USP_LABELSDATA"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Connection = SqlConnectionclass.GetConnection()
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@MODE", "UPDATE")
                    .Parameters.AddWithValue("@CUSTOMERCODE", txtboxUnit.Text.Trim.ToString())
                    .Parameters.AddWithValue("@INVOICENO", txt_InvoiceNo.Text.Trim.ToString())
                    .Parameters.AddWithValue("@VEDNOR_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@IPADDRESS", gstrIpaddressWinSck)
                    .Parameters.AddWithValue("@DATASTRING", label)
                    .ExecuteNonQuery()
                    .Dispose()
                End With

                'BarcodeLib.Barcode.Linear.Linear barcode = new BarcodeLib.Barcode.Linear.Linear();
                Dim barcode As New BarcodeLib.Barcode.Linear.Linear()
                Dim imageData() As Byte
                'BarcodeLib.Barcode.Linear.Linear()
                With Me.fspgrid_Contractdtls
                    For intLoopCounter = 1 To .MaxRows

                        VAL_SELECT = Nothing
                        .GetText(ENUMLABELDETAILS.VAL_SELECT, intLoopCounter, VAL_SELECT)
                        If IsNothing(VAL_SELECT) = True Then VAL_SELECT = String.Empty

                        If VAL_SELECT = "1" Then

                            VAL_LABELS = Nothing
                            .GetText(ENUMLABELDETAILS.VAL_LABELS, intLoopCounter, VAL_LABELS)
                            If IsNothing(VAL_LABELS) = True Then VAL_LABELS = String.Empty

                            barcode.Data = VAL_LABELS.ToString()
                            imageData = barcode.drawBarcodeAsBytes()
                            UpdatelabelData(imageData, VAL_LABELS.ToString())

                        End If

                    Next
                End With

               
                ' here report integration starts
                Dim objReport As ReportDocument
                Dim frmReportViewer As New eMProCrystalReportViewer
                'Dim strSelectionFormula As String = String.Empty

                objReport = frmReportViewer.GetReportDocument()
                frmReportViewer.ShowPrintButton = True
                frmReportViewer.ShowZoomButton = True
                frmReportViewer.ReportHeader = "Transfer Invoice Label Printing"

                With objReport
                    .Load(My.Application.Info.DirectoryPath & "\Reports\PrintLabels_InvoiceTransfer.rpt")
                    .DataDefinition.FormulaFields("MasterLabel").Text = "'N'"
                    .DataDefinition.FormulaFields("PrintType").Text = "'P'"
                    .RecordSelectionFormula = "{Auto_ASN_Labels_temp.VendorCode} = '" & gstrUNITID & "' AND {Auto_ASN_Labels_temp.IPADDRESS}='" & gstrIpaddressWinSck & "' AND {Auto_ASN_Labels_temp.UnitCode}='" & txtboxUnit.Text.Trim.ToString() & "' And {Auto_ASN_Labels_temp.INVOICENO} = '" & txt_InvoiceNo.Text.Trim.ToString() & "'"
                End With

                frmReportViewer.Show()

                ' here report integration ends      
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub UpdatelabelData(ByRef imageData As Object, ByRef label As String)
        Try
            Dim sqlCmd As New SqlCommand()

            With sqlCmd
                .CommandText = "USP_LABELSDATA"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "UPDATEIMAGE")
                .Parameters.AddWithValue("@CUSTOMERCODE", txtboxUnit.Text.Trim.ToString())
                .Parameters.AddWithValue("@INVOICENO", txt_InvoiceNo.Text.Trim.ToString())
                .Parameters.AddWithValue("@VEDNOR_CODE", gstrUNITID)
                .Parameters.AddWithValue("@IPADDRESS", gstrIpaddressWinSck)
                .Parameters.AddWithValue("@DATASTRING", label)
                .Parameters.AddWithValue("@Image", imageData)
                .ExecuteNonQuery()
                .Dispose()
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub chk_Select_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Select.CheckedChanged
        Try
            If chk_Select.Checked = True Then
                With Me.fspgrid_Contractdtls
                    For intLoopCounter = 0 To .MaxRows
                        .SetText(ENUMLABELDETAILS.VAL_SELECT, intLoopCounter + 1, "1")
                    Next
                End With
            End If
            If chk_Select.Checked = False Then
                With Me.fspgrid_Contractdtls
                    For intLoopCounter = 0 To .MaxRows
                        .SetText(ENUMLABELDETAILS.VAL_SELECT, intLoopCounter + 1, "0")
                    Next
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class