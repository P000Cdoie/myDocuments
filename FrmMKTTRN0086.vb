'-------------------------------------------------------------------------------
'(C) 2001 MIND, ALL RIGHTS RESERVED
'FILE NAME      -   FRMMKTTRN0086.VB
'CREATED BY     -   VINOD SINGH
'CREATION DATE  -   20 JAN 2015
'ISSUE ID       -   10736222 - EMPRO - CT2 - ARE3 FUNCTIONALITY
'DESCRIPTION    -   GENERATES NEW ARE-3 NO AGAINST CT2 INVOICE NUMBER
'-------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Excel = Microsoft.Office.Interop.Excel





Public Class FrmMKTTRN0086

#Region "Declarations"

    Dim mintFormtag As Integer
    Private Enum AREReport
        Preview = 1
        PrintToPrinter
    End Enum

#End Region

#Region "Form and Controls Events"

    Private Sub FrmMKTTRN0086_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub FrmMKTTRN0086_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            mdifrmMain.RemoveFormNameFromWindowList = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = False
            Me.Dispose()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FrmMKTTRN0086_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If MsgBox("Do you want to close the screen?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, ResolveResString(100)) = MsgBoxResult.No Then
                e.Cancel = True
                Return
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FrmMKTTRN0086_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            MdiParent = prjMPower.mdifrmMain
            mintFormtag = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            Call FitToClient(Me, GrpMain, ctlHeader, PnlBtn)
            optNewARE.Checked = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub optNewARE_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optNewARE.CheckedChanged
        Try
            If optNewARE.Checked = True Then
                ClearFields()
                EnableDisableFields(True)
                txtARENo.Enabled = False
                txtARENo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdAREhelp.Enabled = False
                btnSave.Enabled = True
                btnExport.Enabled = False
                btnPreview.Enabled = False
                btnPrint.Enabled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub optReprint_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optReprint.CheckedChanged
        Try
            If optReprint.Checked = True Then
                ClearFields()
                EnableDisableFields(False)
                txtARENo.Enabled = True
                txtARENo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdAREhelp.Enabled = True
                btnSave.Enabled = False
                btnExport.Enabled = True
                btnPreview.Enabled = True
                btnPrint.Enabled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub cmdInvHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInvHelp.Click
        Dim strInvNo() As String
        Dim strQry As String
        Try
            If optNewARE.Checked Then
                strQry = " SELECT INVOICE_NO,INVOICE_DATE FROM VW_PENDING_INV_FOR_ARE_GENERATION WHERE UNIT_CODE='" & gstrUNITID & "'"
                strInvNo = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                If IsNothing(strInvNo) = True Then Exit Sub
                If strInvNo.GetUpperBound(0) <> -1 Then
                    If (Len(strInvNo(0)) >= 1) And strInvNo(0) = "0" Then
                        MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        txtInvNo.Text = strInvNo(0)
                        txtInvNo.Focus()
                        SendKeys.Send("{Tab}")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub txtInvNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInvNo.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        Try
            If Ascii = 13 Then
                SendKeys.Send("{Tab}")
                e.KeyChar = Chr(Ascii)
            Else
                e.KeyChar = Chr(validateKey(txtInvNo.Text, txtInvNo.Text.Trim.Length, Ascii, 12, 0))
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvNo.TextChanged
        Try

            If optNewARE.Checked Then
                txtInvDate.Text = ""
                txtCustomer.Text = ""
                txtCT2No.Text = ""
                txtGrossWT.Text = ""
                txtMarksOnPkg.Text = ""
                txtRemarks.Text = ""
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtInvNo.Validating
        Dim StrSQL As String = String.Empty
        Dim dt As DataTable
        Try
            If txtInvNo.Text.Length > 0 Then
                StrSQL = "SELECT * FROM VW_PENDING_INV_FOR_ARE_GENERATION WHERE UNIT_CODE='" & gstrUNITID & "' AND INVOICE_NO = " & Val(txtInvNo.Text.Trim) & " "
                dt = SqlConnectionclass.GetDataTable(StrSQL)
                If dt.Rows.Count = 0 Then
                    MsgBox("Invalid Invoice No.!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtInvNo.Text = ""
                    e.Cancel = True
                    txtInvNo.Focus()
                    Return
                Else
                    txtInvDate.Text = Convert.ToDateTime(dt.Rows(0)("Invoice_date")).ToString(gstrDateFormat)
                    txtCustomer.Text = dt.Rows(0)("Customer").ToString
                    txtCT2No.Text = GetCT2No(Val(txtInvNo.Text))
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CmdAREhelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdAREhelp.Click
        Dim strARENo() As String
        Dim strQry As String
        Try
            If optReprint.Checked Then
                strQry = "SELECT CAST (ARE_NO AS VARCHAR(10)) ARE_NO, CONVERT(VARCHAR(15),ARE_DATE,103) ARE_DATE FROM ARE3_MST where UNIT_CODE ='" & gstrUNITID & "'"
                strARENo = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                If IsNothing(strARENo) = True Then Exit Sub
                If strARENo.GetUpperBound(0) <> -1 Then
                    If (Len(strARENo(0)) >= 1) And strARENo(0) = "0" Then
                        MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        txtARENo.Text = strARENo(0)
                        FillARE3Details(txtARENo.Text.Trim)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub txtARENo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtARENo.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        Try
            If Ascii = 13 Then
                SendKeys.Send("{Tab}")
                e.KeyChar = Chr(Ascii)
            Else
                e.KeyChar = Chr(validateKey(txtARENo.Text, txtARENo.Text.Trim.Length, Ascii, 12, 0))
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtARENo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtARENo.Validating
        Dim strQry As String
        Dim dt As New DataTable
        Try
            If optReprint.Checked = False Then Exit Sub
            If txtARENo.Text.Trim.Length > 0 Then
                strQry = "SELECT ARE_NO,ARE_DATE FROM ARE3_MST where UNIT_CODE ='" & gstrUNITID & "' and ARE_NO=" & Val(txtARENo.Text) & ""
                dt = SqlConnectionclass.GetDataTable(strQry)
                If dt.Rows.Count > 0 Then
                    FillARE3Details(Val(txtARENo.Text))
                Else
                    MsgBox("Invalid ARE3 No !", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtARENo.Text = ""
                    txtARENo.Focus()
                    e.Cancel = True
                    Return
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If IsNothing(dt) = False Then dt.Dispose()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If ValidateSave() = False Then Exit Sub
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "USP_SAVE_ARE"
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@INV_NO", SqlDbType.VarChar, 10).Value = txtInvNo.Text
                    .Parameters.Add("@INV_DATE", SqlDbType.DateTime).Value = getDateForDB(txtInvDate.Text)
                    .Parameters.Add("@CT2_NO", SqlDbType.VarChar, 500).Value = txtCT2No.Text.Trim
                    .Parameters.Add("@GROSS_WT", SqlDbType.Money).Value = Val(txtGrossWT.Text.Trim)
                    .Parameters.Add("@MARK_ON_PKG", SqlDbType.VarChar, 200).Value = txtMarksOnPkg.Text.Trim
                    .Parameters.Add("@REMARKS", SqlDbType.VarChar, 200).Value = txtRemarks.Text.Trim
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 16).Value = mP_User
                    .Parameters.Add("@ARE_NO", SqlDbType.Int).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    If .Parameters("@ARE_NO").Value > 0 Then
                        txtARENo.Text = .Parameters("@ARE_NO").Value.ToString
                        MsgBox("Record Saved. New ARE-3 No. Generated is - " & txtARENo.Text & "", MsgBoxStyle.Information, ResolveResString(100))
                        btnSave.Enabled = False
                        btnPreview.Enabled = True
                        btnPrint.Enabled = True
                        btnExport.Enabled = True
                    Else
                        MsgBox("Record could not be saved", MsgBoxStyle.Critical, ResolveResString(100))
                    End If
                End With
            End Using
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Try
            optNewARE.Checked = True
            ClearFields()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtGrossWT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGrossWT.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        Try
            If Ascii = 13 Then
                SendKeys.Send("{Tab}")
                e.KeyChar = Chr(Ascii)
            Else
                e.KeyChar = Chr(validateKey(txtGrossWT.Text, txtGrossWT.Text.Trim.Length, Ascii, 8, 2))
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click

        Dim sqlCmd As New SqlCommand
        Dim sqlCon As SqlConnection
        Dim sqlAdp As SqlDataAdapter
        Dim dtHdr As New DataTable
        Dim dtDtl As New DataTable
        Dim ds As New DataSet

        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        Dim intExlRow, intItemCount As Integer
        Dim dblValueTotal, dblDutyTotal, dblqtyofgoods, dblpkg As Double

        Try
            If ValidateReportPrinting() = False Then Exit Sub
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            sqlCon = SqlConnectionclass.GetConnection
            With sqlCmd
                .CommandText = "USP_ARE_REPORT"
                .CommandType = CommandType.StoredProcedure
                .Connection = sqlCon
                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@ARE_NO", SqlDbType.Int).Value = txtARENo.Text
                .Parameters.Add("@IP_ADDR", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
            End With

            sqlAdp = New SqlDataAdapter(sqlCmd)
            sqlAdp.Fill(ds, "T")
            dtHdr = ds.Tables(0)
            dtDtl = ds.Tables(1)

            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add
            oExcel.Windows(1).DisplayGridlines = True
            oSheet = oBook.Worksheets(1)
            oSheet.Name = "ARE-3 FORM"

            oSheet.Range("M2").Value = "FORM ARE-3"
            oSheet.Range("M2").Cells.HorizontalAlignment = Excel.Constants.xlLeft
            oSheet.Range("M2").Font.Bold = True

            oSheet.Range("K3").Value = "(See Rule20(2) of the Central Excise(No.2),Rules,2001)"
            oSheet.Range("K3").Cells.HorizontalAlignment = Excel.Constants.xlLeft
            oSheet.Range("K3").Font.Bold = True

            oSheet.Range("A4").Value = "ANNEXURE"
            oSheet.Range("A4").Cells.HorizontalAlignment = Excel.Constants.xlLeft

            oSheet.Range("A5").Value = "ARE-3 No"
            oSheet.Range("A5").Cells.HorizontalAlignment = Excel.Constants.xlLeft

            oSheet.Range("A6").Value = "Date:"
            oSheet.Range("A6").Cells.HorizontalAlignment = Excel.Constants.xlLeft

            oSheet.Range("C5").Value = dtHdr.Rows(0)("ARE_NO")
            oSheet.Range("C5").Cells.HorizontalAlignment = Excel.Constants.xlLeft

            oSheet.Range("C6").Value = Convert.ToDateTime(dtHdr.Rows(0)("ARE_DATE")).ToString(gstrDateFormat)
            oSheet.Range("C6").Cells.HorizontalAlignment = Excel.Constants.xlLeft

            oSheet.Range("V5").Value = "Original/Duplicate/Triplicate/Quadruplicate"
            oSheet.Range("V5").Cells.HorizontalAlignment = Excel.Constants.xlRight

            oSheet.Range("V6").Value = "Range : " + Convert.ToString(dtHdr.Rows(0)("CUST_RANGE"))
            oSheet.Range("V6").Cells.HorizontalAlignment = Excel.Constants.xlRight

            oSheet.Range("V7").Value = "Division : " + Convert.ToString(dtHdr.Rows(0)("CUST_DIVISION"))
            oSheet.Range("V7").Cells.HorizontalAlignment = Excel.Constants.xlRight

            oSheet.Range("A8").Value = "Application for removal of excisable goods from a factory or a warehouse to another warehouse"
            oSheet.Range("A8").Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("A8:V8").Merge()

            oSheet.Range("A9").Value = "(Also called A.R.E 3 for Export Warehousing)"
            oSheet.Range("A9").Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("A9:V9").Merge()

            Dim strDeclaration As String
            If (gstrUNITID = "MST" Or gstrUNITID = "MSB") Then
                strDeclaration = "I/We holder(s)  of  Central Excise Registration No. @UNIT_ECCNO  have undertaken to  remove the undermentioned goods from the factory(Bonded Warehouise)at Motherson Automotive Technologies & Engineering (A Division of Motherson Sumi Systems Ltd.)Plot No.11 Sector -1 ,PhaseII ,Talekuppe ,Manchanaykanahalli,Bidadi Industial Area,Ramanagara, -562109,Karnataka,India.To Factory/Warehouse of M/s.Toyota Kirloskar Motor Private Limited,Plot No.01,Bidadi Industrial Area,Ramnagar District ,bangalore,Karnataka -562109 . (In Both Vendor and Customer address , only customer address has been reflected )."
            Else
                strDeclaration = "I/We holder(s)  of  Central Excise Registration No. @UNIT_ECCNO  have undertaken to  remove the under-mentioned goods from the factory/warehouse to the warehouse at @CUST_NAME,@CUST_ADDR in Deputy Commissioner of Central Excise,Large tax payer Unit, @OFFC_CITY of  @CUST_NAME holders of Central Excise Registration No @CUST_ECCNO"
            End If


            strDeclaration = strDeclaration.Replace("@UNIT_ECCNO", Convert.ToString(dtHdr.Rows(0)("UNIT_ECCNO")))
            strDeclaration = strDeclaration.Replace("@CUST_NAME", Convert.ToString(dtHdr.Rows(0)("CUST_NAME")))
            strDeclaration = strDeclaration.Replace("@CUST_ADDR", Convert.ToString(dtHdr.Rows(0)("CUST_ADDR")))
            strDeclaration = strDeclaration.Replace("@CUST_ECCNO", Convert.ToString(dtHdr.Rows(0)("CUST_ECCNO")))
            strDeclaration = strDeclaration.Replace("@OFFC_CITY", Convert.ToString(dtHdr.Rows(0)("OFFC_CITY")))

            oSheet.Range("A10").Value = strDeclaration
            oSheet.Range("A10").Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("A10:V10").Merge()
            oSheet.Range("A10:V10").WrapText = True
            oSheet.Range("A10").Cells.HorizontalAlignment = Excel.Constants.xlLeft

            oSheet.Range("A11").Value = "Number and date of entry in ware-house register" + vbCrLf + "(1)"
            oSheet.Range("A11:B11").Merge()
            oSheet.Range("A11:B11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("C11").Value = "Description of Goods" + vbCrLf + "(2)"
            oSheet.Range("C11:D11").Merge()
            oSheet.Range("C11:D11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("E11").Value = "No. and description of package" + vbCrLf + "(3)"
            oSheet.Range("E11:F11").Merge()
            oSheet.Range("E11:F11").BorderAround(, Excel.XlBorderWeight.xlThin)


            oSheet.Range("G11").Value = "Gross weight of packages" + vbCrLf + "(4)"
            oSheet.Range("G11:H11").Merge()
            oSheet.Range("G11:H11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("I11").Value = "Marks and numbers on packages" + vbCrLf + "(5)"
            oSheet.Range("I11:J11").Merge()
            oSheet.Range("I11:J11").BorderAround(, Excel.XlBorderWeight.xlThin)


            oSheet.Range("K11").Value = "Quantity Of goods" + vbCrLf + "(6)"
            oSheet.Range("K11:L11").Merge()
            oSheet.Range("K11:L11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("M11").Value = "Date of first warehousing" + vbCrLf + "(7)"
            oSheet.Range("M11:N11").Merge()
            oSheet.Range("M11:N11").BorderAround(, Excel.XlBorderWeight.xlThin)


            oSheet.Range("O11").Value = "Value" + vbCrLf + "(8)"
            oSheet.Range("O11:P11").Merge()
            oSheet.Range("O11:P11").BorderAround(, Excel.XlBorderWeight.xlThin)


            oSheet.Range("Q11").Value = "DUTY " + Convert.ToString(dtHdr.Rows(0)("DUTY_PERC")) + vbCrLf + "(9)"
            oSheet.Range("Q11:R11").Merge()
            oSheet.Range("Q11:R11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("S11").Value = "No. & date of invoice(s) for removal fo goods" + vbCrLf + "(10)"
            oSheet.Range("S11:T11").Merge()
            oSheet.Range("S11:T11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("U11").Value = "Manner of Transport" + vbCrLf + "(11)"
            oSheet.Range("U11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("V11").Value = "Remarks" + vbCrLf + "(12)"
            oSheet.Range("V11").BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("A11:V11").Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("A11:V11").WrapText = True
            oSheet.Range("A11:V11").Font.Bold = True
            intExlRow = 11
            For Each row As DataRow In dtDtl.Rows
                intExlRow += 1
                intItemCount += 1
                If (gstrUNITID = "MST" Or gstrUNITID = "MSB") Then
                    oSheet.Range("A" & intExlRow).Value = row("CT2_NO")
                    oSheet.Range("A" & intExlRow & ":B" & intExlRow).Merge()
                    oSheet.Range("A" & intExlRow & ":B" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
                Else
                    oSheet.Range("A" & intExlRow).Value = intItemCount
                    oSheet.Range("A" & intExlRow & ":B" & intExlRow).Merge()
                    oSheet.Range("A" & intExlRow & ":B" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
                End If
                
                oSheet.Range("C" & intExlRow).Value = row("CUST_DRGNO") + vbCrLf + row("DRG_DESC")
                oSheet.Range("C" & intExlRow & ":D" & intExlRow).Merge()
                oSheet.Range("C" & intExlRow & ":D" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                dblpkg += Val(row("NOS_ON_PKG"))
                oSheet.Range("E" & intExlRow).Value = Convert.ToString(row("NOS_ON_PKG"))
                oSheet.Range("E" & intExlRow & ":F" & intExlRow).Merge()
                oSheet.Range("E" & intExlRow & ":F" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                If Not IsDBNull(row("GROSS_WT")) Then
                    oSheet.Range("G" & intExlRow).Value = row("GROSS_WT")
                End If
                oSheet.Range("G" & intExlRow & ":H" & intExlRow).Merge()
                oSheet.Range("G" & intExlRow & ":H" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                If Not IsDBNull(row("MARK_ON_PKG")) Then
                    oSheet.Range("I" & intExlRow).Value = row("MARK_ON_PKG")
                End If
                oSheet.Range("I" & intExlRow).Value = row("MARK_ON_PKG")
                oSheet.Range("I" & intExlRow & ":J" & intExlRow).Merge()
                oSheet.Range("I" & intExlRow & ":J" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                dblqtyofgoods += Val(row("QTY"))
                oSheet.Range("K" & intExlRow).Value = row("QTY")
                oSheet.Range("K" & intExlRow & ":L" & intExlRow).Merge()
                oSheet.Range("K" & intExlRow & ":L" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                oSheet.Range("M" & intExlRow).Value = ""
                oSheet.Range("M" & intExlRow & ":N" & intExlRow).Merge()
                oSheet.Range("M" & intExlRow & ":N" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                dblValueTotal += Val(row("VALUE"))
                oSheet.Range("O" & intExlRow & ":P" & intExlRow).Merge()
                oSheet.Range("O" & intExlRow & ":P" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                dblDutyTotal += Val(row("DUTY_VALUE"))
                oSheet.Range("Q" & intExlRow & ":R" & intExlRow).Merge()
                oSheet.Range("Q" & intExlRow & ":R" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                If Not IsDBNull(row("INV_NO")) Then
                    oSheet.Range("S" & intExlRow).Value = Convert.ToString(row("INV_NO")) + " dt " + Convert.ToDateTime(row("INV_DATE")).ToString(gstrDateFormat)
                End If

                oSheet.Range("S" & intExlRow & ":T" & intExlRow).Merge()
                oSheet.Range("S" & intExlRow & ":T" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                If Not IsDBNull(row("TRANSPORT")) Then
                    oSheet.Range("U" & intExlRow).Value = row("TRANSPORT")
                End If
                oSheet.Range("U" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                oSheet.Range("V" & intExlRow).Value = row("REMARKS")
                oSheet.Range("V" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

                oSheet.Rows(intExlRow & ":" & intExlRow).RowHeight = 40
                oSheet.Range("A" & intExlRow & ":V" & intExlRow).WrapText = True
                oSheet.Range("A" & intExlRow & ":V" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter

            Next

            oSheet.Range("O12").Value = dblValueTotal
            oSheet.Range("Q12").Value = dblDutyTotal

            oSheet.Range("O12:O" & intExlRow).Merge()
            oSheet.Range("O12:O" & intExlRow).VerticalAlignment = Excel.Constants.xlCenter

            oSheet.Range("Q12:Q" & intExlRow).Merge()
            oSheet.Range("Q12:Q" & intExlRow).VerticalAlignment = Excel.Constants.xlCenter

            oSheet.Range("G12:G" & intExlRow).Merge()
            oSheet.Range("G12:G" & intExlRow).VerticalAlignment = Excel.Constants.xlCenter

            oSheet.Range("I12:I" & intExlRow).Merge()
            oSheet.Range("I12:I" & intExlRow).VerticalAlignment = Excel.Constants.xlCenter

            oSheet.Range("S12:S" & intExlRow).Merge()
            oSheet.Range("S12:S" & intExlRow).VerticalAlignment = Excel.Constants.xlCenter

            oSheet.Range("U12:U" & intExlRow).Merge()
            oSheet.Range("U12:U" & intExlRow).VerticalAlignment = Excel.Constants.xlCenter

            oSheet.Range("V12:V" & intExlRow).Merge()
            oSheet.Range("V12:V" & intExlRow).VerticalAlignment = Excel.Constants.xlCenter

            intExlRow += 1
            oSheet.Range("A" & intExlRow & ":B" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("A" & intExlRow & ":B" & intExlRow).Merge()

            oSheet.Range("C" & intExlRow & ":D" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("C" & intExlRow & ":D" & intExlRow).Merge()

            oSheet.Range("E" & intExlRow & ":F" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("E" & intExlRow & ":F" & intExlRow).Merge()

            oSheet.Range("G" & intExlRow & ":H" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("G" & intExlRow & ":H" & intExlRow).Merge()

            oSheet.Range("I" & intExlRow & ":J" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("I" & intExlRow & ":J" & intExlRow).Merge()

            oSheet.Range("K" & intExlRow & ":L" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("K" & intExlRow & ":L" & intExlRow).Merge()

            oSheet.Range("M" & intExlRow & ":N" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("M" & intExlRow & ":N" & intExlRow).Merge()

            oSheet.Range("O" & intExlRow & ":P" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("O" & intExlRow & ":P" & intExlRow).Merge()

            oSheet.Range("Q" & intExlRow & ":R" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("Q" & intExlRow & ":R" & intExlRow).Merge()

            oSheet.Range("S" & intExlRow & ":T" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("S" & intExlRow & ":T" & intExlRow).Merge()

            oSheet.Range("U" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("U" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("V" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("C" & intExlRow).Value = "TOTAL"
            oSheet.Range("E" & intExlRow).Value = dblpkg
            oSheet.Range("K" & intExlRow).Value = dblqtyofgoods
            oSheet.Range("O" & intExlRow).Value = dblValueTotal
            oSheet.Range("Q" & intExlRow).Value = dblDutyTotal

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "I/We hereby declare the above particulars to be true."

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "Place:"
            If (gstrUNITID = "MST" Or gstrUNITID = "MSB") Then
                oSheet.Range("B" & intExlRow).Value = " BIDADI" 
            Else
                oSheet.Range("B" & intExlRow).Value = gstr_WRK_ADDRESS2
                oSheet.Range("R" & intExlRow & ":V" & intExlRow + 1).WrapText = True
                oSheet.Range("R" & intExlRow & ":V" & intExlRow + 1).Merge()

            End If
            
            oSheet.Range("R" & intExlRow).Value = "For " + gstrCOMPANY

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "Date:"
            oSheet.Range("B" & intExlRow).Value = GetServerDate.ToString("dd-MMM-yyyy")
            oSheet.Range("B" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlLeft
            oSheet.Range("B" & intExlRow & ":C" & intExlRow).Merge()

            If Not (gstrUNITID = "MST" Or gstrUNITID = "MSB") Then
                intExlRow += 1
                oSheet.Range("G" & intExlRow).Value = "(SELF REMOVAL PROCEDURE CLEARANCE)"

            End If
            
            intExlRow += 2
            oSheet.Range("V" & intExlRow).Value = "Authorised Signatory"

            oSheet.Rows("10:10").RowHeight = 59.25
            oSheet.Rows("11:11").RowHeight = 70.25
            oSheet.Columns("B:B").ColumnWidth = 5.29
            oSheet.Columns("D:D").ColumnWidth = 20.14
            oSheet.Columns("F:F").ColumnWidth = 2.57
            oSheet.Columns("H:H").ColumnWidth = 3.43
            oSheet.Columns("J:J").ColumnWidth = 3.29
            oSheet.Columns("L:L").ColumnWidth = 3.71
            oSheet.Columns("N:N").ColumnWidth = 3.43
            oSheet.Columns("P:P").ColumnWidth = 6.86
            oSheet.Columns("R:R").ColumnWidth = 11.14
            oSheet.Columns("T:T").ColumnWidth = 10.57
            oSheet.Columns("U:U").ColumnWidth = 8.43
            oSheet.Columns("V:V").ColumnWidth = 22.57

            intExlRow += 1
            oSheet.HPageBreaks.Add(oSheet.Range("A" & intExlRow))
            
            intExlRow += 4
            'PRINT CERTIFICATE OF WAREHOUSING BY THE CONSIGNEE
            oSheet.Range("A" & intExlRow).Value = "Certificate of warehousing by the consignee"
            oSheet.Range("A" & intExlRow & ":V" & intExlRow).Merge()
            oSheet.Range("A" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "(on original and duplicate)"
            oSheet.Range("A" & intExlRow & ":V" & intExlRow).Merge()
            oSheet.Range("A" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter

            intExlRow += 1
            strDeclaration = "I/We hereby certify that the consignment arrived at…………………………on………………………that the goods conform in all respects to the description given overleaf except for the following discrepancies, and that they have been warehoused under Entry No……………………………of the register maintained in the ware house."
            oSheet.Range("A" & intExlRow).Value = strDeclaration
            oSheet.Range("A" & intExlRow & ":V" & intExlRow).Merge()
            oSheet.Range("A" & intExlRow & ":V" & intExlRow).WrapText = True
            oSheet.Range(intExlRow & ":" & intExlRow).RowHeight = 40

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "Particulars of discrepancies"
            oSheet.Range("A" & intExlRow & ":V" & intExlRow).Merge()
            oSheet.Range("A" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("A" & intExlRow & ":V" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "No.and description of packages not received" + vbCrLf + "(1)"
            oSheet.Range("A" & intExlRow & ":F" & intExlRow).Merge()
            oSheet.Range("A" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("A" & intExlRow & ":F" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range("A" & intExlRow & ":F" & intExlRow).Merge()

            oSheet.Range("G" & intExlRow).Value = "Quanitity short received" + vbCrLf + "(2)"
            oSheet.Range("G" & intExlRow & ":K" & intExlRow).Merge()
            oSheet.Range("G" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("G" & intExlRow & ":K" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("L" & intExlRow).Value = "Duty payable on the shortage" + vbCrLf + "(3)"
            oSheet.Range("L" & intExlRow & ":Q" & intExlRow).Merge()
            oSheet.Range("L" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("L" & intExlRow & ":Q" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)

            oSheet.Range("R" & intExlRow).Value = "Duty payable on the shortage" + vbCrLf + "(4)"
            oSheet.Range("R" & intExlRow & ":V" & intExlRow).Merge()
            oSheet.Range("R" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlCenter
            oSheet.Range("R" & intExlRow & ":V" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
            oSheet.Range(intExlRow & ":" & intExlRow).RowHeight = 40

            For intI As Integer = 1 To 5
                intExlRow += 1
                oSheet.Range("A" & intExlRow & ":F" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
                oSheet.Range("A" & intExlRow & ":F" & intExlRow).Merge()
                oSheet.Range("G" & intExlRow & ":K" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
                oSheet.Range("G" & intExlRow & ":K" & intExlRow).Merge()
                oSheet.Range("L" & intExlRow & ":Q" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
                oSheet.Range("L" & intExlRow & ":Q" & intExlRow).Merge()
                oSheet.Range("R" & intExlRow & ":V" & intExlRow).BorderAround(, Excel.XlBorderWeight.xlThin)
                oSheet.Range("R" & intExlRow & ":V" & intExlRow).Merge()
            Next

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "Place........."
            oSheet.Range("A" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlLeft

            intExlRow += 1
            oSheet.Range("A" & intExlRow).Value = "Date........."
            oSheet.Range("A" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlLeft

            oSheet.Range("V" & intExlRow).Value = "Signature of consignee(s) or his/their"
            oSheet.Range("V" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlRight

            intExlRow += 1
            oSheet.Range("V" & intExlRow).Value = "Authorized Agent."
            oSheet.Range("V" & intExlRow).Cells.HorizontalAlignment = Excel.Constants.xlRight
            intExlRow += 1


            oSheet.HPageBreaks.Add(oSheet.Range("A" & intExlRow))

            oExcel.ActiveWindow.View = Excel.XlWindowView.xlPageBreakPreview
            oSheet.VPageBreaks(1).DragOff(Excel.XlDirection.xlToRight, 1)

            oExcel.Visible = True
            oBook = Nothing
            oSheet = Nothing
            oExcel = Nothing

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If IsNothing(dtDtl) = False Then dtDtl.Dispose()
            If IsNothing(dtHdr) = False Then dtHdr.Dispose()
            If IsNothing(sqlCmd) = False Then sqlCmd.Dispose()
            If IsNothing(sqlCon) = False Then
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
                sqlCon.Dispose()
            End If
            If IsNothing(sqlAdp) = False Then sqlAdp.Dispose()
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Try
            If ValidateReportPrinting() = False Then Exit Sub
            GenerateAREReport(AREReport.Preview)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If ValidateReportPrinting() = False Then Exit Sub
            GenerateAREReport(AREReport.PrintToPrinter)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "Comman Methods"

    Private Sub GenerateAREReport(ByVal PrintType As AREReport)
        Dim strSelectFormula As String
        Dim ObjRpt As ReportDocument
        Dim frmReportViewer As New eMProCrystalReportViewer
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Try
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_ARE_REPORT"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@ARE_NO", SqlDbType.Int).Value = txtARENo.Text
                    .Parameters.Add("@IP_ADDR", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                End With
            End Using
            ObjRpt = frmReportViewer.GetReportDocument()
            frmReportViewer.ShowPrintButton = True
            frmReportViewer.ShowTextSearchButton = True
            frmReportViewer.ShowZoomButton = True
            frmReportViewer.Zoom = 100
            strSelectFormula = ""
            With ObjRpt
                .Load(My.Application.Info.DirectoryPath & "\reports\rptAREPrinting.rpt")
                .DataDefinition.FormulaFields("Place").Text = "'" & gstr_WRK_ADDRESS2 & "'"
                .DataDefinition.FormulaFields("Unit").Text = "'For " & gstrCOMPANY & " ' "
                .RecordSelectionFormula = "{TMP_ARE_REPORT_DTL.UNIT_CODE}='" & gstrUNITID & "' and {TMP_ARE_REPORT_DTL.IP_ADDR}='" & gstrIpaddressWinSck & "'"
            End With
            If PrintType = AREReport.PrintToPrinter Then
                frmReportViewer.SetReportDocument()
                ObjRpt.PrintToPrinter(1, False, 0, 0)
            Else
                frmReportViewer.Show()
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Function ValidateSave() As Boolean
        Dim strSQL As String = String.Empty
        Try
            If Val(txtInvNo.Text.Trim) = 0 Then
                MsgBox("Please select Invoice No.", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtInvNo.Focus()
                Return False
            End If

            If Val(txtGrossWT.Text.Trim) = 0 Then
                MsgBox("Gross Weight can not be Zero.", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtGrossWT.Focus()
                Return False
            End If

            strSQL = "SELECT TOP 1 1 FROM VW_PENDING_INV_FOR_ARE_GENERATION WHERE UNIT_CODE='" & gstrUNITID & "' AND INVOICE_NO = " & Val(txtInvNo.Text.Trim) & ""
            If IsRecordExists(strSQL) = False Then
                MsgBox("Invalid Invoice No selected.", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtInvNo.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function GetCT2No(ByVal intInvNo As Integer) As String
        Dim dt As New DataTable
        Dim strSQL As String
        Dim CT2No As String = String.Empty
        Try
            strSQL = "SELECT DISTINCT CT2_NO FROM CT2_INVOICE_KNOCKOFF WHERE UNIT_CODE ='" & gstrUNITID & "' AND ACT_INV_NO=" & intInvNo & ""
            dt = SqlConnectionclass.GetDataTable(strSQL)
            If dt.Rows.Count > 0 Then
                For Each row As DataRow In dt.Rows
                    CT2No += row("CT2_NO") + ","
                Next
            End If
            Return CT2No
        Catch ex As Exception
            Throw ex
        Finally
            If IsNothing(dt) = False Then
                dt.Dispose()
            End If
        End Try
    End Function

    Private Sub FillARE3Details(ByVal intARENo As Integer)
        Dim dt As New DataTable
        Try
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_ARE_DETAIL"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@ARE_NO", SqlDbType.Int).Value = intARENo
                    dt = SqlConnectionclass.GetDataTable(sqlcmd)
                    If dt.Rows.Count > 0 Then
                        Me.txtMarksOnPkg.Text = Convert.ToString(dt.Rows(0)("MARK_ON_PKG"))
                        Me.txtCT2No.Text = Convert.ToString(dt.Rows(0)("CT2_NO"))
                        Me.txtCustomer.Text = Convert.ToString(dt.Rows(0)("CUSTOMER"))
                        Me.txtGrossWT.Text = Convert.ToString(dt.Rows(0)("GROSS_WT"))
                        Me.txtInvDate.Text = Convert.ToDateTime(dt.Rows(0)("INV_DATE")).ToString(gstrDateFormat)
                        Me.txtRemarks.Text = Convert.ToString(dt.Rows(0)("REMARKS"))
                        Me.txtInvNo.Text = Convert.ToString(dt.Rows(0)("INV_NO"))
                    Else
                        MsgBox("No Record Found !", MsgBoxStyle.Exclamation, ResolveResString(100))
                    End If
                End With
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If IsNothing(dt) = False Then dt.Dispose()
        End Try
    End Sub

    Private Sub ClearFields()
        Try
            Me.txtMarksOnPkg.Text = ""
            Me.txtARENo.Text = ""
            Me.txtCT2No.Text = ""
            Me.txtCustomer.Text = ""
            Me.txtGrossWT.Text = ""
            Me.txtInvDate.Text = ""
            Me.txtInvNo.Text = ""
            Me.txtRemarks.Text = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub EnableDisableFields(ByVal blnStatus As Boolean)
        Try
            Me.txtMarksOnPkg.Enabled = blnStatus
            Me.txtARENo.Enabled = blnStatus
            Me.txtGrossWT.Enabled = blnStatus
            Me.txtInvDate.Enabled = blnStatus
            Me.txtRemarks.Enabled = blnStatus
            CmdAREhelp.Enabled = blnStatus
            cmdInvHelp.Enabled = blnStatus
            If blnStatus = True Then
                Me.txtMarksOnPkg.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.txtARENo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.txtGrossWT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.txtRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.txtInvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Else
                Me.txtMarksOnPkg.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Me.txtARENo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Me.txtGrossWT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Me.txtRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Me.txtInvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function ValidateReportPrinting() As Boolean
        Try
            If txtARENo.Text.Trim = "" Then
                MsgBox("Please select ARE-3 No. first !", MsgBoxStyle.Information, ResolveResString(100))
                If txtARENo.Enabled = True Then txtARENo.Focus()
                Return False
            End If
            If IsNumeric(txtARENo.Text.Trim) = False Then
                MsgBox("Invalid ARE-3 No. !", MsgBoxStyle.Information, ResolveResString(100))
                If txtARENo.Enabled = True Then txtARENo.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function

#End Region
End Class