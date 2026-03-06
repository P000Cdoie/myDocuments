Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Imaging
Imports System.Drawing.Text
Imports System.Runtime.InteropServices
'----------------------------------------------------------------------------------------------
'COPYRIGHT (C)  -   MIND
'NAME OF MODULE -   FRMMKTTRN0103.FRM
'CREATED BY     -   SHUBHRA VERMA
'CREATED DATE   -   MAR 2018
'REVISED BY     -
'REVISED ON     -   
'----------------------------------------------------------------------------------------------
Public Class FRMMKTTRN0103
    Private qrCodeCutOffDate As String = String.Empty
    Private Enum EnumInv
        ItemCode = 1
        ItemCodeHelp
        HSN_SAC_No
        HSN_SAC_Type
        Quantity
        Rate
        Basic_value
        Assessible_Value
        IGST_Tax_type
        IGST_Tax_Per
        IGST_Tax_Value
        CGST_Tax_type
        CGST_Tax_Per
        CGST_Tax_Value
        SGST_Tax_type
        SGST_Tax_Per
        SGST_Tax_Value
        UTGST_Tax_type
        UTGST_Tax_Per
        UTGST_Tax_Value
        ItemTotal
        Internal_Item_Desc
        Cust_Drgno
        Cust_DrgNo_Desc
    End Enum

    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)

        Dim strDesSql As String
        Dim rsDescription As ClsResultSetDB
        Try

            If pstrFieldName2 = "Customer_Code" Then
                strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "'"
            Else
                strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "'"
            End If

            rsDescription = New ClsResultSetDB
            rsDescription.GetResult(strDesSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsDescription.GetNoRows > 0 Then
                pContrName.Text = rsDescription.GetValue(Trim(pstrFieldName1))
            End If
            rsDescription.ResultSetClose()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub BlankFields()
        Try
            lblCancelledInvoice.Visible = False
            txtChallanNo.Text = String.Empty
            txtEmpCode.Text = String.Empty
            lblEmpName.Text = String.Empty
            txtRemarks.Text = ""

            Me.lblInternalPartDesc.Text = String.Empty
            Me.lblCustPartDesc.Text = String.Empty
            sspr.MaxRows = 0
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub FRMMKTTRN0103_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Try
            If e.KeyChar = "'" Then e.Handled = True
            If e.KeyChar = "" Then
                setMode("View")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub FRMMKTTRN0103_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            Me.Visible = False
            Me.KeyPreview = True
            Me.MdiParent = mdifrmMain
            Me.Icon = mdifrmMain.Icon
            Call FitToClient(Me, CObj(PnlMain), ctlHeader, CObj(grpButtons), 250)

            cmdCancelInv.Text = "Cancel Invoice"
            cmdDelete.Text = "Delete Invoice"

            setMode("View")
            qrCodeCutOffDate = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT EMPLOYEE_INVOICE_QR_CODE_CUTOFF_DATE FROM SALES_PARAMETER (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "'"))
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Information, ResolveResString(100))
        Finally
            Me.Visible = True
        End Try
    End Sub

#Region "GRID FORMATTING"
    Private Sub SetPRGridHeading()
        Try
            With Me.sspr

                .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
                .MaxRows = 0
                .MaxCols = EnumInv.Cust_DrgNo_Desc
                .RowHeaderCols = 0
                .Row = 0
                .set_RowHeight(0, 400)
                .ColHeaderRows = 2

                '-- Item Code-------

                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.ItemCode
                .Text = "Item Code"
                .set_ColWidth(EnumInv.ItemCode, 2000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Item Code"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.ItemCode, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.ItemCodeHelp
                .Text = "Help"
                .set_ColWidth(EnumInv.ItemCodeHelp, 500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Help"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.ItemCodeHelp, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.HSN_SAC_No
                .Text = "HSN/" + vbCrLf + "SAC" + vbCrLf + "No"
                .set_ColWidth(EnumInv.HSN_SAC_No, 700)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "HSN/" + vbCrLf + "SAC" + vbCrLf + "No"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.HSN_SAC_No, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                ''- HSN/SAC Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.HSN_SAC_Type
                .Text = "HSN/" + vbCrLf + "SAC" + vbCrLf + "Type"
                .set_ColWidth(EnumInv.HSN_SAC_Type, 700)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "HSN/" + vbCrLf + "SAC" + vbCrLf + "Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.HSN_SAC_Type, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                ''- Quantity
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Quantity
                .Text = "Quantity"
                .set_ColWidth(EnumInv.Quantity, 900)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Quantity"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Quantity, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                ''- Rate
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Rate
                .Text = "Rate"
                .set_ColWidth(EnumInv.Rate, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Rate"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Rate, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                ''- Basic Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Basic_value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Basic Value"
                .set_ColWidth(EnumInv.Basic_value, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.LightPink

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Basic Value"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Basic_value, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .TypeButtonColor = Color.LightPink

                ''- Assessible Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Assessible_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Assessible Value"
                .set_ColWidth(EnumInv.Assessible_Value, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.LightGreen

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Assessible Value"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Assessible_Value, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .TypeButtonColor = Color.LightGreen


                ''- IGST----- IGST TAX TYPE
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.IGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST"
                .set_ColWidth(EnumInv.IGST_Tax_type, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.IGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.SkyBlue


                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "IGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                ''- IGST Tax %
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.IGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST"
                .set_ColWidth(EnumInv.IGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "IGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                ''- IGST Tax Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.IGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST"
                .set_ColWidth(EnumInv.IGST_Tax_Value, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                ''- CGST-----CGST TAX Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST"
                .set_ColWidth(EnumInv.CGST_Tax_type, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .BackColor = Color.Yellow
                .AddCellSpan(EnumInv.CGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "CGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- CGST Tax %
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST"
                .set_ColWidth(EnumInv.CGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "CGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- CGST TAx Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST"
                .set_ColWidth(EnumInv.CGST_Tax_Value, 1100)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                '- SGST--SGST Tax Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.SGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST"
                .set_ColWidth(EnumInv.SGST_Tax_type, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue
                .AddCellSpan(EnumInv.SGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "SGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- SGST Tax Per
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.SGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST"
                .set_ColWidth(EnumInv.SGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "SGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- SGST Tax Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.SGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST"
                .set_ColWidth(EnumInv.SGST_Tax_Value, 1100)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                ''- UTGST----UTGST Tax Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.UTGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST"
                .set_ColWidth(EnumInv.UTGST_Tax_type, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .BackColor = Color.Yellow
                .AddCellSpan(EnumInv.UTGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "UTGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- UTGST Tax %
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.UTGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST"
                .set_ColWidth(EnumInv.UTGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "UTGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- UTGST Tax Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.UTGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST"
                .set_ColWidth(EnumInv.UTGST_Tax_Value, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                ''- ITEM Total
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.ItemTotal
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Item Total"
                .set_ColWidth(EnumInv.ItemTotal, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.Green

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Item Total"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.ItemTotal, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .TypeButtonColor = Color.Green

                ''- Internal Item Code Description.
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Internal_Item_Desc
                .Text = "Internal Item Desc"
                .set_ColWidth(EnumInv.Internal_Item_Desc, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .ColHidden = True

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Internal Item Desc"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Internal_Item_Desc, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .ColHidden = True

                ''- Customer Drawing No .
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Cust_Drgno
                .Text = "Customer DrgNo"
                .set_ColWidth(EnumInv.Cust_Drgno, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Customer DrgNo"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Cust_Drgno, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .ColHidden = True

                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Cust_DrgNo_Desc
                .Text = "Customer DrgNo Desc"
                .set_ColWidth(EnumInv.Cust_DrgNo_Desc, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Customer DrgNo Desc"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Cust_DrgNo_Desc, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .ColHidden = True

            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try

    End Sub
#End Region

    Private Sub CmdCustCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEmpCodeHelp.Click
        Try
            Dim strHelpString As String

            If CmdChallanNo.Enabled = False Then
                strHelpString = ShowList(1, (txtEmpCode.MaxLength), "", "Employee_Code", "Name", , " ", , "EXEC USP_Employee_Inv_Help '" & gstrUNITID & "','Employee','',''")
                If strHelpString = "-1" Then
                    Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Else
                    txtEmpCode.Text = strHelpString
                    lblEmpName.Text = SqlConnectionclass.ExecuteScalar("select Name from employee_mst where unit_code = '" & gstrUNITID & "' and employee_code = '" & txtEmpCode.Text & "' ")
                    lblAddress.Text = SqlConnectionclass.ExecuteScalar("select isnull(Address_1,'') + ' ' + isnull(Address_2,'') + ' ' + isnull(city,'') from employee_mst where unit_code = '" & gstrUNITID & "' and employee_code = '" & txtEmpCode.Text & "' ")
                    Call AddRow()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try

    End Sub

    Private Function GetDocumentNo(ByVal Qselect As String, ByVal HelpFor As String) As String

        Dim Result As String = ""
        Dim strHelp() As String = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, Qselect)

        Try
            If UBound(strHelp) = -1 Then Exit Function
            If Not IsNothing(strHelp) AndAlso strHelp.Length = 0 Then
                Result = "No record found."
                Return Result
            ElseIf String.IsNullOrEmpty(strHelp(1)) Then
                Result = "No record found."
                Return Result
            Else
                If HelpFor = "EMP_No" Then
                    If IsNothing(strHelp(0)) Or IsNothing(strHelp(1)) Or IsNothing(strHelp(2)) Or IsNothing(strHelp(3)) Or IsNothing(strHelp(4)) Then
                        Result = "No record found."
                        Return Result
                    End If
                    Result = strHelp(0).ToString + "~" + strHelp(1).ToString + "~" + strHelp(2).ToString + "~" + strHelp(3).ToString + "~" + strHelp(4).ToString '+ "~" + strHelp(5).ToString + "~" + strHelp(6).ToString
                End If
                Return Result
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Function

    Private Sub SetFromState()
        Dim strSQL As String = ""

        Try
            strSQL = "select GST_STATECODE from GEN_UNITMASTER WHERE Unt_CodeID = '" & gstrUNITID & "'"
            lblFromState.Text = SqlConnectionclass.ExecuteScalar(strSQL)

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub Fill_Grid(ByVal Mode As String)
        Try
            Dim Qselect As String = String.Empty
            If Mode = "ADD" Then
                Qselect = " SELECT A.ITEM_CODE,A.DESCRIPTION,A.ITEM_CODE Cust_Item_Code,A.DESCRIPTION CUST_DRG_DESC,A.HSN_SAC_CODE AS HSNNO,A.HSN_SAC  AS HSN_TYPE,ISNULL(A.STANDARD_RATE,0) AS RATE, " & _
                 " 1 AS QUANTITY,CAST( ISNULL(A.STANDARD_RATE,0)* 1 AS NUMERIC(18,4)) AS BASIC_VALUE,CAST( ISNULL(A.STANDARD_RATE,0)* 1 AS NUMERIC(18,4)) AS Assessible_AMOUNT,0 AS IGST_TAX_TYPE,0 AS IGST_TAX_PER,0 ,0 CGSTTXRT_TYPE," & _
                             " 0 CGST_PERCENT , 0 CGST_AMT ,0 IGSTTXRT_TYPE, 0 IGST_PERCENT,0 IGST_AMT ,0 UTGSTTXRT_TYPE,0 UTGST_PERCENT,0 UTGST_AMT ,0 SGSTTXRT_TYPE,0 SGST_PERCENT,0 SGST_AMT," & _
                             " 0 ITEM_VALUE, 0 ADVANCE,'' REMARKS,0 INSURANCE,0 FRIEGHT_AMOUNT,'' Remarks" & _
                 " FROM ITEM_MST(NOLOCK) A " & _
                 " WHERE A.UNIT_CODE='" & gstrUNITID & "'" & _
                 " AND A.STATUS = 'A'"
            End If

            If Mode.ToUpper = "VIEW" Then
                Qselect = "	   SELECT DISTINCT B.Item_Code,C.Description,B.Cust_Item_Code,B.Cust_Item_Desc,HSNSACCODE AS HSNNO,ISHSNORSAC  AS HSN_TYPE,ISNULL(B.RATE,0) AS RATE" & _
                             " ,B.Sales_Quantity AS QUANTITY,B.Basic_Amount AS BASIC_VALUE,B.Accessible_amount,B.IGSTTXRT_TYPE AS IGST_TAX_TYPE,B.IGST_PERCENT AS IGST_TAX_PER,B.IGST_AMT ,B.CGSTTXRT_TYPE," & _
                             " B.CGST_PERCENT ,B.CGST_AMT ,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT ,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT ,SGSTTXRT_TYPE,SGST_PERCENT,SGST_AMT," & _
                             " B.ITEM_VALUE,A.FRIEGHT_AMOUNT,A.Remarks,A.From_State,A.To_State, a.Bill_Flag, a.Cancel_Flag" & _
                             " FROM EMP_SALESCHALLAN_DTL(NOLOCK) AS A" & _
                             " INNER JOIN EMP_SALES_DTL (NOLOCK) AS B" & _
                             " ON A.UNIT_CODE=B.UNIT_CODE AND A.DOC_NO=B.DOC_NO " & _
                             " INNER JOIN ITEM_MST(NOLOCK) C" & _
                             " ON B.UNIT_CODE=C.UNIT_CODE AND B.ITEM_CODE=C.ITEM_CODE" & _
                             " WHERE B.UNIT_CODE='" & gstrUNITID & "'  AND B.Doc_No='" & txtChallanNo.Text & "'"
            End If

            Dim da As New SqlDataAdapter(Qselect, SqlConnectionclass.GetConnection)
            Dim dt As New DataTable

            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                MessageBox.Show("No Record Found.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                If Mode = "ADD" Then
                    Call PopulateGridData(dt)
                End If
                If Mode = "VIEW" Then
                    Call PopulateGridData(dt)
                    Me.txtRemarks.Text = dt.Rows(0)("REMARKS").ToString
                    If dt.Rows(0)("Cancel_Flag").ToString = "True" Then
                        lblCancelledInvoice.Text = "C A N C E L L E D  I N V O I C E"
                        lblCancelledInvoice.Visible = True
                        lblCancelledInvoice.ForeColor = Color.Red
                    ElseIf dt.Rows(0)("Bill_Flag").ToString = "True" Then
                        lblCancelledInvoice.Text = "I N V O I C E  L O C K E D"
                        lblCancelledInvoice.ForeColor = Color.Green
                        lblCancelledInvoice.Visible = True
                    Else
                        lblCancelledInvoice.Text = ""
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub AddRow()
        Try
            With sspr

                .Row = .MaxRows : .Col = EnumInv.ItemCode
                If (.Text.Length = 0) Then Exit Sub

                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = EnumInv.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeMaxEditLen = 30 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = EnumInv.ItemCodeHelp : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeButtonText = "Help"
                .Row = .MaxRows : .Col = EnumInv.HSN_SAC_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeMaxEditLen = 30 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .Row = .MaxRows : .Col = EnumInv.HSN_SAC_Type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeMaxEditLen = 30 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .Row = .MaxRows : .Col = EnumInv.Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = EnumInv.Rate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = EnumInv.Basic_value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = EnumInv.Assessible_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0.0
                .Row = .MaxRows : .Col = EnumInv.IGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeMaxEditLen = 15 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = EnumInv.IGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = EnumInv.IGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                .Row = .MaxRows : .Col = EnumInv.CGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeMaxEditLen = 15 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = EnumInv.CGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = EnumInv.CGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                .Row = .MaxRows : .Col = EnumInv.SGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeMaxEditLen = 15 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = EnumInv.SGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = EnumInv.SGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeMaxEditLen = 15 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                .Row = .MaxRows : .Col = EnumInv.ItemTotal : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                .Row = .MaxRows : .Col = EnumInv.Internal_Item_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = EnumInv.Cust_DrgNo_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = EnumInv.Cust_Drgno : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .set_RowHeight(.MaxRows, 300)
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub PopulateGridData(ByVal dtRec As DataTable)
        Dim intRow As Integer
        Dim intRowCount As Integer

        Try
            If CmdChallanNo.Enabled = True Then
                sspr.MaxRows = 0
                intRowCount = dtRec.Rows.Count
                lblFromState.Text = dtRec.Rows(0)("From_State").ToString()
                lblToState.Text = dtRec.Rows(0)("To_State").ToString()

                For intRow = 1 To intRowCount

                    AddRow()
                    intRow = sspr.MaxRows

                    With sspr
                        .Row = intRow : .Col = EnumInv.ItemCode : .Text = IIf(dtRec.Rows(intRow - 1)("ITEM_CODE").ToString <> "", dtRec.Rows(intRow - 1)("ITEM_CODE").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.HSN_SAC_No : .Text = IIf(dtRec.Rows(intRow - 1)("HSNNO").ToString <> "", dtRec.Rows(intRow - 1)("HSNNO").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.HSN_SAC_Type : .Text = IIf(dtRec.Rows(intRow - 1)("HSN_TYPE").ToString <> "", dtRec.Rows(intRow - 1)("HSN_TYPE").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.Quantity : .Text = IIf(dtRec.Rows(intRow - 1)("QUANTITY").ToString <> "0", dtRec.Rows(intRow - 1)("QUANTITY").ToString, 0)
                        .Row = intRow : .Col = EnumInv.Rate : .Text = IIf(dtRec.Rows(intRow - 1)("RATE").ToString <> "0", dtRec.Rows(intRow - 1)("RATE").ToString, 0)
                        .Row = intRow : .Col = EnumInv.Basic_value : .Text = IIf(dtRec.Rows(intRow - 1)("BASIC_VALUE").ToString <> "0", dtRec.Rows(intRow - 1)("BASIC_VALUE").ToString, 0)
                        .Row = intRow : .Col = EnumInv.Assessible_Value : .Text = dtRec.Rows(intRow - 1)("Accessible_AMOUNT").ToString
                        .Row = intRow : .Col = EnumInv.IGST_Tax_type : .Text = IIf(dtRec.Rows(intRow - 1)("IGST_TAX_TYPE").ToString <> "", dtRec.Rows(intRow - 1)("IGST_TAX_TYPE").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.IGST_Tax_Per : .Text = dtRec.Rows(intRow - 1)("IGST_TAX_PER").ToString
                        .Row = intRow : .Col = EnumInv.IGST_Tax_Value : .Text = dtRec.Rows(intRow - 1)("IGST_AMT").ToString
                        .Row = intRow : .Col = EnumInv.CGST_Tax_type : .Text = IIf(dtRec.Rows(intRow - 1)("CGSTTXRT_TYPE").ToString <> "", dtRec.Rows(intRow - 1)("CGSTTXRT_TYPE").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.CGST_Tax_Per : .Text = IIf(dtRec.Rows(intRow - 1)("CGST_PERCENT").ToString <> "", dtRec.Rows(intRow - 1)("CGST_PERCENT").ToString, 0.0)
                        .Row = intRow : .Col = EnumInv.CGST_Tax_Value : .Text = dtRec.Rows(intRow - 1)("CGST_AMT").ToString
                        .Row = intRow : .Col = EnumInv.SGST_Tax_type : .Text = IIf(dtRec.Rows(intRow - 1)("SGSTTXRT_TYPE").ToString <> "", dtRec.Rows(intRow - 1)("SGSTTXRT_TYPE").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.SGST_Tax_Per : .Text = IIf(dtRec.Rows(intRow - 1)("SGST_PERCENT").ToString <> "", dtRec.Rows(intRow - 1)("SGST_PERCENT").ToString, 0.0)
                        .Row = intRow : .Col = EnumInv.SGST_Tax_Value : .Text = IIf(dtRec.Rows(intRow - 1)("SGST_AMT").ToString <> "", dtRec.Rows(intRow - 1)("SGST_AMT").ToString, 0.0)
                        .Row = intRow : .Col = EnumInv.UTGST_Tax_type : .Text = IIf(dtRec.Rows(intRow - 1)("UTGSTTXRT_TYPE").ToString <> "", dtRec.Rows(intRow - 1)("UTGSTTXRT_TYPE").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.UTGST_Tax_Per : .Text = dtRec.Rows(intRow - 1)("UTGST_PERCENT").ToString
                        .Row = intRow : .Col = EnumInv.UTGST_Tax_Value : .Text = dtRec.Rows(intRow - 1)("UTGST_AMT").ToString
                        .Row = intRow : .Col = EnumInv.ItemTotal : .Text = IIf(dtRec.Rows(intRow - 1)("ITEM_VALUE").ToString <> "", dtRec.Rows(intRow - 1)("ITEM_VALUE").ToString, 0)
                        .Row = intRow : .Col = EnumInv.Internal_Item_Desc : .Text = IIf(dtRec.Rows(intRow - 1)("DESCRIPTION").ToString <> "", dtRec.Rows(intRow - 1)("DESCRIPTION").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.Cust_DrgNo_Desc : .Text = IIf(dtRec.Rows(intRow - 1)("Cust_Item_Desc").ToString <> "", dtRec.Rows(intRow - 1)("Cust_Item_Desc").ToString, String.Empty)
                        .Row = intRow : .Col = EnumInv.Cust_Drgno : .Text = IIf(dtRec.Rows(intRow - 1)("CUST_ITEM_CODE").ToString <> "", dtRec.Rows(intRow - 1)("CUST_ITEM_CODE").ToString, String.Empty)
                    End With
                Next

            Else

                intRow = sspr.ActiveRow

                With sspr
                    .Row = intRow : .Col = EnumInv.ItemCode : .Text = IIf(dtRec.Rows(0)("ITEM_CODE").ToString <> "", dtRec.Rows(0)("ITEM_CODE").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.HSN_SAC_No : .Text = IIf(dtRec.Rows(0)("HSNNO").ToString <> "", dtRec.Rows(0)("HSNNO").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.HSN_SAC_Type : .Text = IIf(dtRec.Rows(0)("HSN_TYPE").ToString <> "", dtRec.Rows(0)("HSN_TYPE").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.Quantity : .Text = IIf(dtRec.Rows(0)("QUANTITY").ToString <> "0", dtRec.Rows(0)("QUANTITY").ToString, 0)
                    .Row = intRow : .Col = EnumInv.Rate : .Text = IIf(dtRec.Rows(0)("RATE").ToString <> "0", dtRec.Rows(0)("RATE").ToString, 0)
                    .Row = intRow : .Col = EnumInv.Basic_value : .Text = IIf(dtRec.Rows(0)("BASIC_VALUE").ToString <> "0", dtRec.Rows(0)("BASIC_VALUE").ToString, 0)
                    .Row = intRow : .Col = EnumInv.Assessible_Value : .Text = dtRec.Rows(0)("Accessible_AMOUNT").ToString
                    .Row = intRow : .Col = EnumInv.IGST_Tax_type : .Text = IIf(dtRec.Rows(0)("IGST_TAX_TYPE").ToString <> "", dtRec.Rows(0)("IGST_TAX_TYPE").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.IGST_Tax_Per : .Text = dtRec.Rows(0)("IGST_TAX_PER").ToString
                    .Row = intRow : .Col = EnumInv.IGST_Tax_Value : .Text = dtRec.Rows(0)("IGST_AMT").ToString
                    .Row = intRow : .Col = EnumInv.CGST_Tax_type : .Text = IIf(dtRec.Rows(0)("CGSTTXRT_TYPE").ToString <> "", dtRec.Rows(0)("CGSTTXRT_TYPE").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.CGST_Tax_Per : .Text = IIf(dtRec.Rows(0)("CGST_PERCENT").ToString <> "", dtRec.Rows(0)("CGST_PERCENT").ToString, 0.0)
                    .Row = intRow : .Col = EnumInv.CGST_Tax_Value : .Text = dtRec.Rows(0)("CGST_AMT").ToString
                    .Row = intRow : .Col = EnumInv.SGST_Tax_type : .Text = IIf(dtRec.Rows(0)("SGSTTXRT_TYPE").ToString <> "", dtRec.Rows(0)("SGSTTXRT_TYPE").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.SGST_Tax_Per : .Text = IIf(dtRec.Rows(0)("SGST_PERCENT").ToString <> "", dtRec.Rows(0)("SGST_PERCENT").ToString, 0.0)
                    .Row = intRow : .Col = EnumInv.SGST_Tax_Value : .Text = IIf(dtRec.Rows(0)("SGST_AMT").ToString <> "", dtRec.Rows(0)("SGST_AMT").ToString, 0.0)
                    .Row = intRow : .Col = EnumInv.UTGST_Tax_type : .Text = IIf(dtRec.Rows(0)("UTGSTTXRT_TYPE").ToString <> "", dtRec.Rows(0)("UTGSTTXRT_TYPE").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.UTGST_Tax_Per : .Text = dtRec.Rows(0)("UTGST_PERCENT").ToString
                    .Row = intRow : .Col = EnumInv.UTGST_Tax_Value : .Text = dtRec.Rows(0)("UTGST_AMT").ToString
                    .Row = intRow : .Col = EnumInv.ItemTotal : .Text = IIf(dtRec.Rows(0)("ITEM_VALUE").ToString <> "", dtRec.Rows(0)("ITEM_VALUE").ToString, 0)
                    .Row = intRow : .Col = EnumInv.Internal_Item_Desc : .Text = IIf(dtRec.Rows(0)("DESCRIPTION").ToString <> "", dtRec.Rows(0)("DESCRIPTION").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.Cust_DrgNo_Desc : .Text = IIf(dtRec.Rows(0)("Cust_Item_Desc").ToString <> "", dtRec.Rows(0)("Cust_Item_Desc").ToString, String.Empty)
                    .Row = intRow : .Col = EnumInv.Cust_Drgno : .Text = IIf(dtRec.Rows(0)("CUST_ITEM_CODE").ToString <> "", dtRec.Rows(0)("CUST_ITEM_CODE").ToString, String.Empty)
                End With

            End If
            If CmdChallanNo.Enabled = False Then
                Calculate_Grid_Values(intRow)
            End If

            If CmdChallanNo.Enabled = True Then
                With sspr
                    .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols
                    .BlockMode = True : .Lock = True : .BlockMode = False
                End With
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

#Region "CALCULATE TAXATION AMOUNT"

    Private Sub Calculate_Grid_Values(ByVal intRow As Integer)
        Dim BasicVal As Object
        Dim ItemTotal As Object
        Dim AssessibleValue As Object

        Dim IGST_Tax_Per As Object
        Dim CGST_Tax_Per As Object
        Dim SGST_Tax_Per As Object
        Dim UTGST_Tax_Per As Object

        Dim IGST_Tax_Val As Object
        Dim CGST_Tax_Val As Object
        Dim SGST_Tax_Val As Object
        Dim UTGST_Tax_Val As Object
        Dim Quantity As Object
        Dim Rate As Object
        Dim Basic_value As Object

        Try
            With sspr
                .Row = intRow
                Quantity = Nothing
                Rate = Nothing
                Basic_value = Nothing
                BasicVal = Nothing
                ItemTotal = Nothing
                AssessibleValue = Nothing

                IGST_Tax_Per = Nothing
                CGST_Tax_Per = Nothing
                SGST_Tax_Per = Nothing
                UTGST_Tax_Per = Nothing

                IGST_Tax_Val = Nothing
                CGST_Tax_Val = Nothing
                SGST_Tax_Val = Nothing
                UTGST_Tax_Val = Nothing

                .Col = EnumInv.Quantity
                Quantity = Val(.Text)

                .Col = EnumInv.Rate
                Rate = CDec(Val(.Text))

                .Col = EnumInv.Basic_value
                .Text = CDec(Quantity) * CDec(Rate)

                .GetText(EnumInv.Assessible_Value, .Row, BasicVal)
                If IsNothing(BasicVal) = True Then BasicVal = 0

                .Col = EnumInv.IGST_Tax_Per
                .GetText(EnumInv.IGST_Tax_Per, .Row, IGST_Tax_Per)
                If IsNothing(IGST_Tax_Per) = True Then IGST_Tax_Per = 0
                .Text = IGST_Tax_Per
                .Col = EnumInv.IGST_Tax_Value

                IGST_Tax_Val = BasicVal * CDec(IGST_Tax_Per / 100)
                If IsNothing(IGST_Tax_Val) = True Then IGST_Tax_Val = 0
                .Text = IGST_Tax_Val

                .Col = EnumInv.CGST_Tax_Per
                .GetText(EnumInv.CGST_Tax_Per, .Row, CGST_Tax_Per)
                If IsNothing(CGST_Tax_Per) = True Then CGST_Tax_Per = 0

                .Col = EnumInv.CGST_Tax_Value
                CGST_Tax_Val = BasicVal * CDec(CGST_Tax_Per / 100)
                If IsNothing(CGST_Tax_Val) = True Then CGST_Tax_Val = 0
                .Text = CGST_Tax_Val

                .Col = EnumInv.SGST_Tax_Per
                .GetText(EnumInv.SGST_Tax_Per, .Row, SGST_Tax_Per)
                If IsNothing(SGST_Tax_Per) = True Then SGST_Tax_Per = 0

                .Col = EnumInv.SGST_Tax_Value
                SGST_Tax_Val = BasicVal * CDec(SGST_Tax_Per / 100)
                If IsNothing(SGST_Tax_Val) = True Then SGST_Tax_Val = 0
                .Text = SGST_Tax_Val

                .Col = EnumInv.UTGST_Tax_Per
                .GetText(EnumInv.UTGST_Tax_Per, .Row, UTGST_Tax_Per)
                If IsNothing(UTGST_Tax_Per) = True Then UTGST_Tax_Per = 0

                .Col = EnumInv.UTGST_Tax_Value
                UTGST_Tax_Val = BasicVal * CDec(UTGST_Tax_Per / 100)
                If IsNothing(UTGST_Tax_Val) = True Then UTGST_Tax_Val = 0
                .Text = UTGST_Tax_Val

                ItemTotal = CDec(IGST_Tax_Val) + CDec(CGST_Tax_Val) + CDec(SGST_Tax_Val) + CDec(UTGST_Tax_Val) + (CDec(Quantity) * CDec(Rate))
                .Col = EnumInv.ItemTotal
                If IsNothing(ItemTotal) = True Then ItemTotal = 0
                .Text = CDec(ItemTotal)

            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
#End Region

    Private Sub sspr_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles sspr.ButtonClicked
        Dim strHelpString As String
        Dim strSQL As String = String.Empty
        Dim i As Integer
        Dim dt As New DataTable
        Dim strItems As String = String.Empty

        Try

            If e.col = EnumInv.ItemCodeHelp Then
                If lblToState.Text = "" Then
                    MessageBox.Show("Select To State", ResolveResString(100), MessageBoxButtons.OK)
                    Exit Sub
                End If

                If sspr.MaxRows > 1 Then
                    For i = 1 To sspr.MaxRows
                        If i <> sspr.ActiveRow Then
                            sspr.Row = i : sspr.Col = EnumInv.ItemCode
                            strItems = strItems + sspr.Text + ","
                        End If
                    Next
                End If
                If strItems.Length > 0 Then
                    strItems = Mid(strItems, 1, Len(strItems) - 1)
                End If

                strSQL = "EXEC USP_Employee_Inv_Help '" & gstrUNITID & "','ITEM','" & lblFromState.Text & "','" & lblToState.Text & "','" & strItems & "'"

                strHelpString = ShowList(1, (txtEmpCode.MaxLength), "", "Item_code", "Description", "", "", "Item Help", strSQL, , , , )
                If strHelpString = "-1" Then
                    MessageBox.Show("No record found:" + vbCrLf + vbCrLf + "Check the following-" + vbCrLf + "Item Status must be active." + vbCrLf + "Item HSN must be linked with selected state HSN.", ResolveResString(100), MessageBoxButtons.OK)
                Else
                    With sspr
                        strSQL = "SELECT A.ITEM_CODE,A.DESCRIPTION,A.ITEM_CODE CUST_ITEM_CODE,A.DESCRIPTION CUST_ITEM_DESC,A.HSN_SAC_CODE AS HSNNO,A.HSN_SAC  AS HSN_TYPE,ISNULL(A.STANDARD_RATE,0) AS RATE,"
                        strSQL = strSQL + " 1 AS QUANTITY,CAST( ISNULL(A.STANDARD_RATE,0)* 1 AS NUMERIC(18,4)) AS BASIC_VALUE,CAST( ISNULL(A.STANDARD_RATE,0)* 1 AS NUMERIC(18,4)) AS Accessible_AMOUNT,"
                        strSQL = strSQL + " 0 AS IGST_TAX_TYPE,0 AS IGST_TAX_PER,0 ,0 CGSTTXRT_TYPE, 0 CGST_PERCENT , 0 CGST_AMT ,0 IGSTTXRT_TYPE, 0 IGST_PERCENT,0 IGST_AMT ,"
                        strSQL = strSQL + " 0 UTGSTTXRT_TYPE,0 UTGST_PERCENT,0 UTGST_AMT ,0 SGSTTXRT_TYPE,0 SGST_PERCENT,0 SGST_AMT,"
                        strSQL = strSQL + " 0 ITEM_VALUE, 0 ADVANCE,'' REMARKS,0 INSURANCE,0 FRIEGHT_AMOUNT,'' Remarks FROM ITEM_MST(NOLOCK) A  "
                        strSQL = strSQL + " WHERE A.UNIT_CODE='" & gstrUNITID & "' AND A.item_code = '" & strHelpString & "'"

                        Dim da As New SqlDataAdapter(strSQL, SqlConnectionclass.GetConnection)

                        da.Fill(dt)
                        Call PopulateGridData(dt)
                        FillTaxDetails()
                        Calculate_Grid_Values(e.row)
                        AddRow()

                    End With
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub sspr_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles sspr.EditChange
        Dim objRate As Object
        Dim objQty As Object

        Try

            If (e.col = EnumInv.Quantity) Or (e.col = EnumInv.Rate) Then
                With sspr
                    .Row = e.row
                    .Col = EnumInv.Rate : objRate = .Text
                    .Col = EnumInv.Quantity : objQty = .Text
                    .Col = EnumInv.Assessible_Value : .Text = objQty * objRate
                End With
            End If

            If (e.col = EnumInv.Assessible_Value) Or (e.col = EnumInv.Quantity) Or (e.col = EnumInv.Rate) Then
                Calculate_Grid_Values(e.row)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub sspr_ClickEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles sspr.ClickEvent
        Try
            With sspr
                .Row = e.row
                .Col = EnumInv.Internal_Item_Desc
                lblInternalPartDesc.Text = .Text

                .Col = EnumInv.Cust_DrgNo_Desc
                lblCustPartDesc.Text = .Text

            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Function SaveData() As Boolean
        Dim DocumentNo As String = String.Empty
        Dim cmd As SqlCommand = Nothing

        Try
            SaveData = False

            With sspr
                .Row = .MaxRows : .Col = EnumInv.ItemCode
                If .Text.Length = 0 Then .MaxRows = .MaxRows - 1
            End With

            If Me.txtEmpCode.Text.Trim = String.Empty Then
                MsgBox("Please Select The Customer.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.sspr.MaxRows = 0 Then
                MsgBox("Please Select Item To Be Despatched.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            sspr.Row = 1
            sspr.Col = EnumInv.ItemTotal
            If Val(sspr.Text) = 0 Then
                MsgBox("Item Quantity In Invoice Can't Be Zero.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            Dim ValidateItem_Code As Object
            Dim ValidateItem_total As Object
            For intCount As Integer = 1 To sspr.MaxRows
                With sspr
                    .Col = EnumInv.ItemCode
                    ValidateItem_Code = .Text

                    .Col = EnumInv.Quantity
                    If Val(.Text) = 0 Then
                        MessageBox.Show("Quantity Can not be Zero for Item-" & ValidateItem_Code & " at Row No-" & intCount, "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If
                    .Col = EnumInv.ItemTotal
                    ValidateItem_total = Val(.Text)
                    If Val(ValidateItem_total) = 0 Then
                        MessageBox.Show("Item Total Can not be Zero for Item-" & ValidateItem_Code & " at Row No-" & intCount, "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If
                End With
            Next

            SqlConnectionclass.CloseGlobalConnection()
            SqlConnectionclass.OpenGlobalConnection()

            cmd = New System.Data.SqlClient.SqlCommand()
            cmd.Connection = SqlConnectionclass.GetConnection
            cmd.Transaction = cmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)

            Dim QryExecuted As Boolean = False
            If CmdChallanNo.Enabled = False Then
                Try
                    Dim CGST_VAL As Object = 0
                    Dim IGST_VAL As Object = 0
                    Dim SGST_VAL As Object = 0
                    Dim UTGST_VAL As Object = 0
                    Dim ITEM_TOTAL As Object = 0
                    Dim Error_Msg As Object = String.Empty

                    For i As Integer = 1 To sspr.MaxRows
                        With sspr
                            .Row = i
                            .Col = EnumInv.CGST_Tax_Value
                            CGST_VAL = Val(CGST_VAL) + Val(.Text)

                            .Col = EnumInv.IGST_Tax_Value
                            IGST_VAL = Val(IGST_VAL) + Val(.Text)

                            .Col = EnumInv.SGST_Tax_Value
                            SGST_VAL = Val(SGST_VAL) + Val(.Text)

                            .Col = EnumInv.UTGST_Tax_Value
                            UTGST_VAL = Val(UTGST_VAL) + Val(.Text)

                            .Col = EnumInv.ItemTotal
                            ITEM_TOTAL = Val(ITEM_TOTAL) + Val(.Text)

                        End With
                    Next

                    With cmd
                        QryExecuted = False
                        .CommandText = String.Empty
                        .Parameters.Clear()
                        .CommandTimeout = 0
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "USP_EMPLOYEE_INVOICE"

                        .Parameters.Add("@EMP_No", SqlDbType.VarChar, 30).Value = txtEmpCode.Text.Trim.ToString
                        .Parameters.Add("@Emp_Name", SqlDbType.VarChar, 200).Value = lblEmpName.Text.ToString
                        .Parameters.Add("@User_Id", SqlDbType.VarChar, 100).Value = mP_User
                        .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@CGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(CGST_VAL.ToString)
                        .Parameters.Add("@SGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(SGST_VAL.ToString)
                        .Parameters.Add("@IGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(IGST_VAL.ToString)
                        .Parameters.Add("@UTGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(UTGST_VAL.ToString)
                        .Parameters.Add("@FromState", SqlDbType.VarChar, 3).Value = lblFromState.Text
                        .Parameters.Add("@ToState", SqlDbType.VarChar, 3).Value = lblToState.Text
                        .Parameters.Add("@Para", SqlDbType.VarChar, 200).Value = "SAVESALES_DTL"

                        Dim retMessage As SqlParameter = New SqlParameter("@CHALLAN_No", SqlDbType.VarChar, 10)
                        retMessage.Direction = ParameterDirection.Output
                        .Parameters.Add(retMessage)


                        Dim retErrorMessage As SqlParameter = New SqlParameter("@Error_Msg", SqlDbType.VarChar, 8000)
                        retErrorMessage.Direction = ParameterDirection.Output
                        .Parameters.Add(retErrorMessage)

                        cmd.ExecuteNonQuery()
                        DocumentNo = .Parameters("@CHALLAN_No").Value.ToString()
                        Error_Msg = .Parameters("@Error_Msg").Value.ToString()

                        If Not String.IsNullOrEmpty(Error_Msg) Then
                            MessageBox.Show(Error_Msg.ToString, "Empro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            QryExecuted = False
                            cmd.Transaction.Rollback()
                            Exit Function
                        Else
                            If Not String.IsNullOrEmpty(DocumentNo.ToString) Then
                                QryExecuted = True
                            Else
                                cmd.Transaction.Rollback()
                                SaveData = False
                                QryExecuted = False
                                MessageBox.Show("Document Number not generated.", "Empro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Exit Function
                            End If
                        End If

                        If Not String.IsNullOrEmpty(DocumentNo) Then
                            Dim Item_code As Object
                            Dim Itemtotal As Object
                            Dim Rate As Object
                            Dim Cust_Drgno As Object
                            Dim Cust_DrgNo_Desc As Object
                            Dim Basic_Value As Object
                            Dim Assessible_Value As Object
                            Dim Remarks As Object
                            Dim HSN_SAC_NO As Object
                            Dim HSN_SAC_TYPE As Object
                            Dim CGST_Tax_Type As Object
                            Dim CGST_Tax_Per As Object
                            Dim CGST_Tax_Value As Object
                            Dim SGST_Tax_Type As Object
                            Dim SGST_Tax_Per As Object
                            Dim SGST_Tax_Value As Object
                            Dim UTGST_Tax_Type As Object
                            Dim UTGST_Tax_Per As Object
                            Dim UTGST_Tax_Value As Object
                            Dim IGST_Tax_Type As Object
                            Dim IGST_Tax_Per As Object
                            Dim IGST_Tax_Value As Object
                            Dim Quantity As Object

                            For intcounter As Integer = 1 To sspr.MaxRows
                                Item_code = Nothing
                                Itemtotal = Nothing
                                Rate = Nothing
                                Cust_Drgno = Nothing
                                Cust_DrgNo_Desc = Nothing
                                Basic_Value = Nothing
                                Assessible_Value = Nothing
                                Remarks = Nothing
                                HSN_SAC_TYPE = Nothing
                                Quantity = Nothing
                                With sspr
                                    .Row = intcounter

                                    .Col = EnumInv.Quantity
                                    Quantity = Val(.Text)

                                    .Col = EnumInv.ItemCode
                                    Item_code = .Text

                                    .Col = EnumInv.ItemTotal
                                    Itemtotal = Val(.Text)

                                    .Col = EnumInv.Rate
                                    Rate = Val(.Text)

                                    .Col = EnumInv.Cust_Drgno
                                    Cust_Drgno = .Text

                                    .Col = EnumInv.Cust_DrgNo_Desc
                                    Cust_DrgNo_Desc = .Text

                                    .Col = EnumInv.Basic_value
                                    Basic_Value = Val(.Text)

                                    .Col = EnumInv.Assessible_Value
                                    Assessible_Value = Val(.Text)

                                    .Col = EnumInv.HSN_SAC_No
                                    HSN_SAC_NO = .Text

                                    .Col = EnumInv.HSN_SAC_Type
                                    HSN_SAC_TYPE = .Text

                                    .Col = EnumInv.CGST_Tax_type
                                    CGST_Tax_Type = .Text

                                    .Col = EnumInv.CGST_Tax_Per
                                    CGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.CGST_Tax_Value
                                    CGST_Tax_Value = Val(.Text)

                                    .Col = EnumInv.IGST_Tax_type
                                    IGST_Tax_Type = .Text

                                    .Col = EnumInv.IGST_Tax_Per
                                    IGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.IGST_Tax_Value
                                    IGST_Tax_Value = Val(.Text)

                                    .Col = EnumInv.SGST_Tax_type
                                    SGST_Tax_Type = .Text

                                    .Col = EnumInv.SGST_Tax_Per
                                    SGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.SGST_Tax_Value
                                    SGST_Tax_Value = Val(.Text)

                                    .Col = EnumInv.UTGST_Tax_type
                                    UTGST_Tax_Type = .Text

                                    .Col = EnumInv.UTGST_Tax_Per
                                    UTGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.UTGST_Tax_Value
                                    UTGST_Tax_Value = Val(.Text)

                                    With cmd

                                        cmd.CommandText = String.Empty
                                        cmd.Parameters.Clear()
                                        .CommandTimeout = 0
                                        .CommandType = CommandType.StoredProcedure
                                        .CommandText = "USP_EMPLOYEE_INVOICE"

                                        .Parameters.Add("@Doc_No", SqlDbType.VarChar, 10).Value = DocumentNo.ToString.Trim
                                        .Parameters.Add("@Item_Code", SqlDbType.VarChar, 30).Value = Item_code.ToString
                                        .Parameters.Add("@Quantity", SqlDbType.VarChar, 100).Value = Val(Quantity.ToString)
                                        .Parameters.Add("@Rate", SqlDbType.VarChar, 100).Value = Rate.ToString.Trim
                                        .Parameters.Add("@Cust_Drgno", SqlDbType.VarChar, 100).Value = Cust_Drgno.ToString.ToString
                                        .Parameters.Add("@Cust_DrgNo_Desc", SqlDbType.VarChar, 200).Value = Cust_DrgNo_Desc.ToString.ToString
                                        .Parameters.Add("@Basic_Value", SqlDbType.VarChar, 100).Value = Basic_Value.ToString.Trim
                                        .Parameters.Add("@Assessible_Value", SqlDbType.VarChar, 100).Value = Assessible_Value.ToString
                                        .Parameters.Add("@HSN_SAC_NO", SqlDbType.VarChar, 100).Value = HSN_SAC_NO.ToString
                                        .Parameters.Add("@HSN_SAC_Type", SqlDbType.VarChar, 100).Value = HSN_SAC_TYPE.ToString
                                        .Parameters.Add("@CGST_Tax_Type", SqlDbType.VarChar, 200).Value = CGST_Tax_Type.ToString
                                        .Parameters.Add("@CGST_Tax_Per", SqlDbType.VarChar, 100).Value = CGST_Tax_Per.ToString
                                        .Parameters.Add("@CGST_Tax_Value", SqlDbType.VarChar, 100).Value = CGST_Tax_Value.ToString
                                        .Parameters.Add("@SGST_Tax_Type", SqlDbType.VarChar, 200).Value = SGST_Tax_Type.ToString
                                        .Parameters.Add("@SGST_Tax_Per", SqlDbType.VarChar, 100).Value = SGST_Tax_Per.ToString
                                        .Parameters.Add("@SGST_Tax_Value", SqlDbType.VarChar, 100).Value = SGST_Tax_Value.ToString
                                        .Parameters.Add("@UTGST_Tax_Type", SqlDbType.VarChar, 200).Value = UTGST_Tax_Type.ToString
                                        .Parameters.Add("@UTGST_Tax_Per", SqlDbType.VarChar, 100).Value = UTGST_Tax_Per.ToString
                                        .Parameters.Add("@UTGST_Tax_Value", SqlDbType.VarChar, 100).Value = UTGST_Tax_Value.ToString
                                        .Parameters.Add("@IGST_Tax_Type", SqlDbType.VarChar, 200).Value = IGST_Tax_Type.ToString
                                        .Parameters.Add("@IGST_Tax_Per", SqlDbType.VarChar, 100).Value = IGST_Tax_Per.ToString
                                        .Parameters.Add("@IGST_Tax_Value", SqlDbType.VarChar, 100).Value = IGST_Tax_Value.ToString
                                        .Parameters.Add("@Total_Amount", SqlDbType.VarChar, 100).Value = Itemtotal.ToString
                                        .Parameters.Add("@User_Id", SqlDbType.VarChar, 100).Value = mP_User
                                        .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID
                                        .Parameters.Add("@FromState", SqlDbType.VarChar, 3).Value = lblFromState.Text
                                        .Parameters.Add("@ToState", SqlDbType.VarChar, 3).Value = lblToState.Text
                                        .Parameters.Add("@Para", SqlDbType.VarChar, 200).Value = "SAVECHALLAN_DTL"

                                        Dim retError_Message As SqlParameter = New SqlParameter("@Error_Msg", SqlDbType.VarChar, 8000)
                                        retError_Message.Direction = ParameterDirection.Output
                                        .Parameters.Add(retError_Message)

                                        cmd.ExecuteNonQuery()
                                        Error_Msg = .Parameters("@Error_Msg").Value.ToString()
                                        If Not String.IsNullOrEmpty(Error_Msg) Then
                                            cmd.Transaction.Rollback()
                                            MessageBox.Show("4427-Transaction Fail.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                            QryExecuted = False
                                            SaveData = False
                                            Exit Function
                                        Else
                                            QryExecuted = True
                                        End If
                                    End With
                                End With
                            Next
                            If QryExecuted Then
                                cmd.Transaction.Commit()
                                cmd.Dispose()
                                SaveData = True
                                txtChallanNo.Text = DocumentNo.ToString.Trim
                            End If
                        End If

                    End With

                Catch ex As Exception
                    cmd.Transaction.Dispose()
                    RaiseException(ex)
                    Exit Function
                End Try

            End If

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsNothing(cmd) = False Then
                cmd.Dispose()
            End If

        End Try
    End Function

    Private Function LockInv() As Boolean
        Dim DocumentNo As String = String.Empty
        Dim cmd As SqlCommand = Nothing
        Dim Error_Msg As Object = String.Empty

        Try
            LockInv = False

            SqlConnectionclass.CloseGlobalConnection()
            SqlConnectionclass.OpenGlobalConnection()

            cmd = New System.Data.SqlClient.SqlCommand()
            cmd.Connection = SqlConnectionclass.GetConnection
            cmd.Transaction = cmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)

            Dim QryExecuted As Boolean = False

            With cmd
                QryExecuted = False
                .CommandText = String.Empty
                .Parameters.Clear()
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_EMPINV_POSTING"

                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@InvoiceNo", SqlDbType.VarChar, 20).Value = txtChallanNo.Text
                .Parameters.Add("@User", SqlDbType.VarChar, 10).Value = mP_User

                Dim retErrorMessage As SqlParameter = New SqlParameter("@RETMSG", SqlDbType.VarChar, 250)
                retErrorMessage.Direction = ParameterDirection.Output
                .Parameters.Add(retErrorMessage)

                cmd.ExecuteNonQuery()
                Error_Msg = .Parameters("@RETMSG").Value.ToString()

                If Not String.IsNullOrEmpty(Error_Msg) Then
                    MessageBox.Show(Error_Msg.ToString, "Empro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    QryExecuted = False
                    cmd.Transaction.Rollback()
                    Exit Function
                Else
                    QryExecuted = True
                End If

                If QryExecuted Then
                    cmd.Transaction.Commit()
                    cmd.Dispose()
                    LockInv = True
                End If

            End With

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            cmd.Transaction.Rollback()
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsNothing(cmd) = False Then
                cmd.Dispose()
            End If

        End Try
    End Function

    Private Sub CmdChallanNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChallanNo.Click
        Try
            Dim strQselect As String = String.Empty
            Dim EMP_detail As String = String.Empty
            strQselect = " SELECT DISTINCT  A.DOC_NO AS EMP_NO,A.INVOICE_DATE AS EMP_DATE  "
            strQselect = strQselect + " ,A.ACCOUNT_CODE AS CUST_CODE ,A.CUST_NAME  , CASE WHEN A.CANCEL_FLAG=0 THEN 'OPEN' WHEN A.CANCEL_FLAG=1 THEN 'CLOSE' END AS STATUS"
            strQselect = strQselect + " FROM EMP_SALESCHALLAN_DTL(NOLOCK) AS A"
            strQselect = strQselect + " INNER JOIN EMP_SALES_DTL (NOLOCK) AS B"
            strQselect = strQselect + " ON A.UNIT_CODE=B.UNIT_CODE AND A.DOC_NO=B.DOC_NO "
            strQselect = strQselect + " WHERE A.UNIT_CODE='" & gstrUNITID & "' and isDeleted = 0 ORDER BY A.INVOICE_DATE DESC"

            EMP_detail = GetDocumentNo(strQselect, "EMP_No")
            If EMP_detail = Nothing Then
                MessageBox.Show("Operation Cancelled", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Exit Sub
            End If
            Dim Split() As String = EMP_detail.Split("~")
            If Split(0).Contains("No record found.") Then
                MessageBox.Show("No Record Found", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Exit Sub
            End If
            If Not String.IsNullOrEmpty(Split(0)) Then
                txtChallanNo.Text = Split(0).ToString
                If IsDate(Split(1).ToString) Then
                    dtpDate.Text = Split(1).ToString
                Else
                    dtpDate.Text = Now.Date.ToString
                End If
                If Not String.IsNullOrEmpty(Split(2)) Then
                    txtEmpCode.Text = Split(2).ToString
                End If
                If Not String.IsNullOrEmpty(Split(3)) Then
                    lblEmpName.Text = Split(3).ToString
                End If

                Call Fill_Grid("VIEW")

                Me.Group4.Enabled = True
                Me.Group2.Enabled = True

            Else
                txtChallanNo.Text = String.Empty
                Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdToState_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdToState.Click
        Dim strHelpString As String

        Try

            With sspr
                If .MaxRows > 0 Then
                    .Row = 1
                    .Col = EnumInv.ItemCode
                    If .Text.Length > 0 Then
                        MessageBox.Show("After selecting items, State Code cannot be changed!")
                        Exit Sub
                    End If
                End If
            End With


            If CmdChallanNo.Enabled = False Then
                strHelpString = ShowList(1, (txtEmpCode.MaxLength), "", "STATE_TO", "SHORT_STATE_CODE", "", "", "To State Help", " select distinct G.state_to,S.SHORT_STATE_CODE from GST_STATEWISE_TAX_MAPPING G INNER JOIN STATE_MST S ON G.STATE_TO = S.GST_STATE_CODE  where G.STATE_FROM = '" & lblFromState.Text & "' and getdate() between G.TXRT_VALIDFROM and G.TXRT_VALIDTO  and s.active = 1", , , , )
                If strHelpString = "-1" Then
                    Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Else
                    lblToState.Text = strHelpString
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillTaxDetails()
        Dim strSQL As String = String.Empty
        Dim strHSN As String = String.Empty
        Dim strval As String = String.Empty
        Dim oRdr As SqlDataReader

        Try
            With sspr
                .Row = .ActiveRow
                .Col = EnumInv.HSN_SAC_No
                strHSN = .Text

                .Col = EnumInv.Assessible_Value
                strval = .Text

                strSQL = "SELECT TXRT_TYPE,TXRT_PER FROM GST_STATEWISE_TAX_MAPPING WHERE HSNSACCODE = '" & strHSN & "' AND STATE_FROM  = '" & lblFromState.Text & "' AND STATE_TO = '" & lblToState.Text & "' and TXRT_HEAD = 'IGST' AND CONVERT(VARCHAR(12),GETDATE(),106) BETWEEN  TXRT_VALIDFROM  AND TXRT_VALIDTO "
                oRdr = SqlConnectionclass.ExecuteReader(strSQL)
                If oRdr.HasRows Then
                    oRdr.Read()
                    .Col = EnumInv.IGST_Tax_type
                    .Text = oRdr("TXRT_TYPE").ToString

                    .Col = EnumInv.IGST_Tax_Per
                    .Text = oRdr("TXRT_PER").ToString

                    .Col = EnumInv.IGST_Tax_Value
                    .Text = (Val(strval) * Val(oRdr("TXRT_PER").ToString)) / 100
                End If

                strSQL = "SELECT TXRT_TYPE,TXRT_PER FROM GST_STATEWISE_TAX_MAPPING WHERE HSNSACCODE = '" & strHSN & "' AND STATE_FROM  = '" & lblFromState.Text & "' AND STATE_TO = '" & lblToState.Text & "' and TXRT_HEAD = 'CGST' AND CONVERT(VARCHAR(12),GETDATE(),106) BETWEEN  TXRT_VALIDFROM  AND TXRT_VALIDTO "
                oRdr = SqlConnectionclass.ExecuteReader(strSQL)
                If oRdr.HasRows Then
                    oRdr.Read()
                    .Col = EnumInv.CGST_Tax_type
                    .Text = oRdr("TXRT_TYPE").ToString

                    .Col = EnumInv.CGST_Tax_Per
                    .Text = oRdr("TXRT_PER").ToString

                    .Col = EnumInv.CGST_Tax_Value
                    .Text = (Val(strval) * Val(oRdr("TXRT_PER").ToString)) / 100
                End If

                strSQL = "SELECT TXRT_TYPE,TXRT_PER FROM GST_STATEWISE_TAX_MAPPING WHERE HSNSACCODE = '" & strHSN & "' AND STATE_FROM  = '" & lblFromState.Text & "' AND STATE_TO = '" & lblToState.Text & "' and TXRT_HEAD = 'SGST' AND CONVERT(VARCHAR(12),GETDATE(),106) BETWEEN  TXRT_VALIDFROM  AND TXRT_VALIDTO "
                oRdr = SqlConnectionclass.ExecuteReader(strSQL)
                If oRdr.HasRows Then
                    oRdr.Read()
                    .Col = EnumInv.SGST_Tax_type
                    .Text = oRdr("TXRT_TYPE").ToString

                    .Col = EnumInv.SGST_Tax_Per
                    .Text = oRdr("TXRT_PER").ToString

                    .Col = EnumInv.SGST_Tax_Value
                    .Text = (Val(strval) * Val(oRdr("TXRT_PER").ToString)) / 100
                End If

                strSQL = "SELECT TXRT_TYPE,TXRT_PER FROM GST_STATEWISE_TAX_MAPPING WHERE HSNSACCODE = '" & strHSN & "' AND STATE_FROM  = '" & lblFromState.Text & "' AND STATE_TO = '" & lblToState.Text & "' and TXRT_HEAD = 'UTGST' AND CONVERT(VARCHAR(12),GETDATE(),106) BETWEEN TXRT_VALIDFROM  AND TXRT_VALIDTO "
                oRdr = SqlConnectionclass.ExecuteReader(strSQL)
                If oRdr.HasRows Then
                    oRdr.Read()
                    .Col = EnumInv.UTGST_Tax_type
                    .Text = oRdr("TXRT_TYPE").ToString

                    .Col = EnumInv.UTGST_Tax_Per
                    .Text = oRdr("TXRT_PER").ToString

                    .Col = EnumInv.UTGST_Tax_Value
                    .Text = (Val(strval) * Val(oRdr("TXRT_PER").ToString)) / 100
                End If

            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
        Try
            If cmdNew.Text = "New" Then
                setMode("Add")
                cmdNew.Text = "Save"
                lblEscMsg.Text = "Press ESC to cancel operation!"
                lblEscMsg.ForeColor = Color.Blue
            Else
                If SaveData() = True Then
                    MsgBox("Invoice Saved Successfully." + vbCrLf + "Invoice No-" + txtChallanNo.Text.ToString, MsgBoxStyle.Information, ResolveResString(100))
                    setMode("View")
                    lblEscMsg.Text = ""
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdCancelInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelInv.Click
        Dim strSQL As String = String.Empty
        Dim SqlCmd As SqlCommand
        
        Try
            If txtChallanNo.Text.Length = 0 Then Exit Sub

            If txtRemarks.Text = "" Then
                MessageBox.Show("Enter Reason to Cancel the selected invoice.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If

            strSQL = "Select top 1 1 from emp_saleschallan_dtl where unit_code ='" & gstrUNITID & "' and doc_no = '" & txtChallanNo.Text & "' and bill_flag = 0"
            If IsRecordExists(strSQL) Then
                MessageBox.Show("Only locked invoice can be cancelled.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            strSQL = "Select top 1 1 from emp_saleschallan_dtl where unit_code ='" & gstrUNITID & "' and doc_no = '" & txtChallanNo.Text & "' and Cancel_flag = 1"
            If IsRecordExists(strSQL) Then
                MessageBox.Show("This is a cancelled Invoice!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If MessageBox.Show("Are you sure, you want to cancel selected invoice.", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then

                SqlCmd = New SqlCommand
                With SqlCmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .Transaction = SqlCmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "usp_CancelInvoice"
                    .CommandTimeout = 0
                    .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 12).Value = txtChallanNo.Text
                    .Parameters.Add("@User", SqlDbType.VarChar, 20).Value = mP_User
                    .Parameters.Add("@CancelRemarks", SqlDbType.VarChar, 100).Value = txtChallanNo.Text

                    Dim retErrorMessage As SqlParameter = New SqlParameter("@ERR", SqlDbType.VarChar, 250)
                    retErrorMessage.Direction = ParameterDirection.Output
                    .Parameters.Add(retErrorMessage)

                    .ExecuteNonQuery()

                    If .Parameters("@ERR").Value.ToString() = 0 Then
                        SqlCmd.Transaction.Rollback()
                        MessageBox.Show("Error while Invoice Cancellation!", "Empro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        SqlCmd.Transaction.Commit()
                        MessageBox.Show("Invoice Cancelled.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Fill_Grid("View")
                    End If

                End With
                SqlCmd.Dispose()

            End If
        Catch ex As Exception
            SqlCmd.Transaction.Rollback()
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        Finally
            If IsNothing(SqlCmd) = False Then
                SqlCmd.Dispose()
            End If
        End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim strSQL As String = String.Empty

        Try
            If txtChallanNo.Text.Length = 0 Then Exit Sub

            If txtRemarks.Text = "" Then
                MessageBox.Show("Enter Reason to Delete the selected invoice.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If

            strSQL = "Select top 1 1 from emp_saleschallan_dtl where unit_code ='" & gstrUNITID & "' and doc_no = '" & txtChallanNo.Text & "' and bill_flag = 1"
            If IsRecordExists(strSQL) Then
                MessageBox.Show("Invoice Locked. Cannot Delete!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            strSQL = "Select top 1 1 from emp_saleschallan_dtl where unit_code ='" & gstrUNITID & "' and doc_no = '" & txtChallanNo.Text & "' and Cancel_flag = 1"
            If IsRecordExists(strSQL) Then
                MessageBox.Show("This is a cancelled Invoice!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If MessageBox.Show("Are you sure, you want to delete selected invoice.", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                strSQL = "Update emp_saleschallan_dtl set isDeleted = 1, DeleteRemarks = '" & txtRemarks.Text & "' where unit_Code = '" & gstrUNITID & "' and doc_no = '" & txtChallanNo.Text & "'"
                SqlConnectionclass.ExecuteNonQuery(strSQL)

                MessageBox.Show("Invoice Deleted", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtChallanNo.Text = ""
                BlankFields()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdLock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLock.Click
        Dim strSQL As String = String.Empty
        Try

            If txtChallanNo.Text.Length = 0 Then Exit Sub

            strSQL = "Select top 1 1 from emp_saleschallan_dtl where unit_code ='" & gstrUNITID & "' and doc_no = '" & txtChallanNo.Text & "' and Cancel_flag = 1"
            If IsRecordExists(strSQL) Then
                MessageBox.Show("This is a cancelled Invoice!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            strSQL = "Select top 1 1 from emp_saleschallan_dtl where unit_code ='" & gstrUNITID & "' and doc_no = '" & txtChallanNo.Text & "' and Bill_Flag = 1"
            If IsRecordExists(strSQL) Then
                MessageBox.Show("Invoice already Locked!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If MessageBox.Show("Are you sure, you want to lock selected invoice.", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                If LockInv() = True Then
                    MessageBox.Show("Invoice Locked.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Fill_Grid("View")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Try
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub setMode(ByVal Mode As String)
        Try

            cmdFromState.Enabled = False
            SetFromState()

            If Mode = "View" Then
                BlankFields()
                Call SetPRGridHeading()
                Me.Group2.Enabled = False
                Me.Group4.Enabled = False
                Me.CmdChallanNo.Enabled = True
                Me.CmdEmpCodeHelp.Enabled = False
                Me.cmdToState.Enabled = False
                cmdNew.Text = "New"
                lblEscMsg.Text = ""
                lblToState.Text = ""
            End If

            If Mode = "Add" Then
                BlankFields()
                Call SetPRGridHeading()
                Me.Group2.Enabled = True
                Me.Group4.Enabled = True
                Me.CmdChallanNo.Enabled = False
                Me.CmdEmpCodeHelp.Enabled = True
                Me.cmdToState.Enabled = True
                lblToState.Text = ""
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtChallanNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtChallanNo.KeyPress
        Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Dim REPDOC As ReportDocument
        Dim REPVIEWER As eMProCrystalReportViewer
        Dim strReportName As String = ""
        Dim Challan_No As String
        Dim Employee_Code As String = ""
        Dim Emp_Address As String = ""
        Dim Comp_Address As String = ""
        Dim RegNo, Range, phone, Fax, EMail, Division, Commissionerate, Invoice_Rule, strCompMst As String
        Dim rsCompMst As ClsResultSetDB
        Dim SqlCmd As SqlCommand

        Try
            RegNo = ""
            Range = ""
            phone = ""
            Fax = ""
            EMail = ""
            Division = ""
            Commissionerate = ""
            Invoice_Rule = ""
            strCompMst = ""

            If Len(txtChallanNo.Text.Trim) > 0 Then
                Challan_No = txtChallanNo.Text.Trim
            Else
                MessageBox.Show("Please Select Chalan No..!!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If Len(txtEmpCode.Text.Trim) > 0 Then
                Employee_Code = txtEmpCode.Text
            Else
                MessageBox.Show("Employee Code can't be blank..!!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If Len(lblAddress.Text.Trim) > 0 Then
                Emp_Address = lblAddress.Text.Trim
            Else
                Emp_Address = ""
            End If

            B2CQRBarcode(Challan_No)

            rsCompMst = New ClsResultSetDB
            strCompMst = "Select * from Company_Mst WHERE UNIT_CODE = '" & gstrUNITID & "'"
            rsCompMst.GetResult(strCompMst)
            If rsCompMst.GetNoRows = 1 Then
                RegNo = rsCompMst.GetValue("Reg_NO")
                Range = rsCompMst.GetValue("Range_1")
                phone = rsCompMst.GetValue("Phone")
                Fax = rsCompMst.GetValue("Fax")
                EMail = rsCompMst.GetValue("Email")
                Division = rsCompMst.GetValue("Division")
                Commissionerate = rsCompMst.GetValue("Commissionerate")
                Invoice_Rule = rsCompMst.GetValue("Invoice_Rule")
            End If
            rsCompMst.ResultSetClose()

            SqlCmd = New SqlCommand
            With SqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.StoredProcedure
                .CommandText = "PRC_EMPLOYEE_INVOICEPRINTING"
                .CommandTimeout = 0
                .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck.Trim
                .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 10).Value = gstrUNITID.Trim
                .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 10).Value = Challan_No
                .Parameters.Add("@EMP_CODE", SqlDbType.VarChar, 20).Value = Employee_Code.Trim
                SqlCmd.ExecuteNonQuery()
            End With
            SqlCmd.Dispose()
            strReportName = "\Reports\rptEmpInvoice_GST_A4reports" & ".rpt"

            If Not CheckFile(strReportName) Then
                strReportName = My.Application.Info.DirectoryPath & "\Reports\rptEmpInvoice_GST_A4reports.rpt"
            Else
                strReportName = My.Application.Info.DirectoryPath & strReportName
            End If

            REPDOC = New ReportDocument()
            REPVIEWER = New eMProCrystalReportViewer()
            REPDOC = REPVIEWER.GetReportDocument()
            REPDOC.Load(strReportName)
            Comp_Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2

            With REPDOC
                .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
                .DataDefinition.FormulaFields("CompanyAddress").Text = "'" & Comp_Address & "'"
                .DataDefinition.FormulaFields("Phone").Text = "'" & phone & "'"
                .DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
                .DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
                .DataDefinition.FormulaFields("commissionerate").Text = "'" & Commissionerate & "'"
                .DataDefinition.FormulaFields("InvoiceRule").Text = "'" & Invoice_Rule & "'"
                ' .DataDefinition.FormulaFields("Address1").Text = "'" & Emp_Address & "'"
                .DataDefinition.RecordSelectionFormula = "{TMP_EMPLOYEE_INVOICEPRINT.IP_ADDRESS}='" & gstrIpaddressWinSck & "' " & _
                " and {TMP_EMPLOYEE_INVOICEPRINT.DOC_NO}=" & Challan_No & " " & _
                " and {TMP_EMPLOYEE_INVOICEPRINT.ACCOUNT_CODE}='" & Employee_Code & "'  " & _
                " and {TMP_EMPLOYEE_INVOICEPRINT.UNIT_CODE}='" & gstrUNITID & "'"
            End With
            REPVIEWER.Show()
        Catch SqlEx As SqlException
            MsgBox(SqlEx.Message, MsgBoxStyle.Critical, ResolveResString(100))
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Finally
            If IsNothing(SqlCmd) = False Then
                If SqlCmd.Connection.State = ConnectionState.Open Then
                    SqlCmd.Connection.Close()
                End If
                SqlCmd.Dispose()
            End If
        End Try
    End Sub


    Private Sub B2CQRBarcode(ByVal invoiceNo As String)
        Try
            Dim rsGENERATEBARCODE As ClsResultSetDB
            Dim straccountcode As String
            Dim strPrintMethod As String = ""
            Dim strSQL As String = ""
            Dim intTotalNoofSlabs As Integer = 0
            Dim intRow As Short
            Dim strBarcodeMsg As String
            Dim strBarcodeMsg_paratemeter As String
            Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUnitId)
            Dim stimage As ADODB.Stream
            Dim strQuery As String
            Dim Rs As ADODB.Recordset
            Dim pstrPath As String = ""
            Dim blnCROP_QRIMAGE As Boolean = False


            pstrPath = gstrUserMyDocPath
            strSQL = "SELECT TOP 1 1 FROM EMP_SALESCHALLAN_DTL H INNER JOIN EMP_SALES_DTL D ON H.UNIT_CODE=D.UNIT_CODE AND H.DOC_NO=D.DOC_NO WHERE H.UNIT_CODE = '" & gstrUnitId & "' AND H.DOC_NO='" & invoiceNo & "' AND H.BILL_FLAG = 1 AND H.CANCEL_FLAG=0 AND H.ISDELETED=0 AND CONVERT(DATE,H.INVOICE_DATE,103) >= CONVERT(DATE,'" & qrCodeCutOffDate & "',103) "

            If DataExist(strSQL) = True Then
                strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForEmployeeInvoice(gstrUserMyDocPath, invoiceNo, gstrCONNECTIONSTRING)

                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                    Exit Sub
                Else
                    strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                    stimage = New ADODB.Stream
                    stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
                    stimage.Open()
                    pstrPath = pstrPath & "QRBarcodeImgEmpInvoice.wmf"

                    blnCROP_QRIMAGE = True
                    If blnCROP_QRIMAGE = True Then
                        Dim bmp As New Bitmap(pstrPath)
                        Dim picturebox1 As New PictureBox
                        picturebox1.Image = ImageTrim(bmp)
                        picturebox1.Image.Save(pstrPath)
                        picturebox1 = Nothing
                    End If

                    stimage.LoadFromFile(pstrPath)

                    strQuery = "SELECT B2C_QR_CODE FROM EMP_SALESCHALLAN_DTL WHERE UNIT_CODE = '" & gstrUnitId & "' AND Doc_No=" & invoiceNo

                    Rs = New ADODB.Recordset
                    Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                    If Not (Rs.EOF And Rs.BOF) Then
                        Rs.Fields("B2C_QR_CODE").Value = stimage.Read
                        Rs.Update()
                    End If

                    Rs.Update()
                    Rs.Close()
                    Rs = Nothing


                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Function ImageTrim(ByVal img As Bitmap) As Bitmap
        'get image data
        Dim bd As BitmapData = img.LockBits(New Rectangle(Point.Empty, img.Size), ImageLockMode.[ReadOnly], PixelFormat.Format32bppArgb)
        Dim rgbValues As Integer() = New Integer(img.Height * img.Width - 1) {}
        Marshal.Copy(bd.Scan0, rgbValues, 0, rgbValues.Length)
        img.UnlockBits(bd)


        '#Region "determine bounds"
        Dim left As Integer = bd.Width
        Dim top As Integer = bd.Height
        Dim right As Integer = 0
        Dim bottom As Integer = 0

        'determine top
        For i As Integer = 0 To rgbValues.Length - 1
            Dim color As Integer = rgbValues(i) And &HFFFFFF
            If color <> &HFFFFFF Then
                Dim r As Integer = i / bd.Width
                Dim c As Integer = i Mod bd.Width

                If left > c Then
                    left = c
                End If
                If right < c Then
                    right = c
                End If
                bottom = r
                top = r
                Exit For
            End If
        Next

        'determine bottom
        For i As Integer = rgbValues.Length - 1 To 0 Step -1
            Dim color As Integer = rgbValues(i) And &HFFFFFF
            If color <> &HFFFFFF Then
                Dim r As Integer = i / bd.Width
                Dim c As Integer = i Mod bd.Width

                If left > c Then
                    left = c
                End If
                If right < c Then
                    right = c
                End If
                bottom = r
                Exit For
            End If
        Next

        If bottom > top Then
            For r As Integer = top + 1 To bottom - 1
                'determine left
                For c As Integer = 0 To left - 1
                    Dim color As Integer = rgbValues(r * bd.Width + c) And &HFFFFFF
                    If color <> &HFFFFFF Then
                        If left > c Then
                            left = c
                            Exit For
                        End If
                    End If
                Next

                'determine right
                For c As Integer = bd.Width - 1 To right + 1 Step -1
                    Dim color As Integer = rgbValues(r * bd.Width + c) And &HFFFFFF
                    If color <> &HFFFFFF Then
                        If right < c Then
                            right = c
                            Exit For
                        End If
                    End If
                Next
            Next
        End If

        Dim width As Integer = right - left + 1
        Dim height As Integer = bottom - top + 1
        '#End Region

        'copy image data
        Dim imgData As Integer() = New Integer(width * height - 1) {}
        For r As Integer = top To bottom
            Array.Copy(rgbValues, r * bd.Width + left, imgData, (r - top) * width, width)
        Next

        'create new image
        Dim newImage As New Bitmap(width, height, PixelFormat.Format32bppArgb)
        Dim nbd As BitmapData = newImage.LockBits(New Rectangle(0, 0, width, height), ImageLockMode.[WriteOnly], PixelFormat.Format32bppArgb)
        Marshal.Copy(imgData, 0, nbd.Scan0, imgData.Length)
        newImage.UnlockBits(nbd)

        ImageTrim = newImage
    End Function
End Class