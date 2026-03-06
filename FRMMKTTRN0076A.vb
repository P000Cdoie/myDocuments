Imports System.Data.SqlClient
Imports System.Text

'*****************************************************************************************
'   Revised By:       Rajeev Gupta
'   Revised On:       13 Jun 2013
'   Issue ID  :       10406132 - Schindengen Changes for RG23D and PO Report is not Opening in view Mode
'*****************************************************************************************
    ''Revised By:       Saurav Kumar
    ''Revised On:       04 Oct 2013
    ''Issue ID  :       10462231 - eMpro ISuite Changes
    '***********************************************************************************************************************************

Public Class FRMMKTTRN0076A
    Public strCustPartNo, strInternalPartNo As String, strInternalPartDesc As String, strCustomerPartDesc As String, strCurrentStockQty As String
    Public ParentFormOperationMode As String = String.Empty
    Public strInvoiceNo As String = String.Empty
    Public blnBillFlag As Boolean = False
    Private Enum GridGrin_ENUM
        GRINNO = 1
        GRINDATE
        GRINQTY
        REMAININGQTY
        KNOCKOFFQTY
        EXCISEPERPIECE
        PAGENUMBER
        AEDPERPIECE
    End Enum
    Private Sub PopulateGrins()
        Dim strSql As String
        Dim Sqlcmd As New SqlCommand
        Dim Dr As SqlDataReader
        Dim SQLCon As SqlConnection
        Dim builder As New StringBuilder
        Dim dblSalesQty As Double
        GridGrin.MaxRows = 0
        SQLCon = SqlConnectionclass.GetConnection()
        Sqlcmd.Connection = SQLCon
        Sqlcmd.CommandType = CommandType.Text
        GridGrin.BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
        Try
            If blnBillFlag = True Then
                builder.AppendLine("SELECT  SLNO,GRINNO GRIN_NO,CONVERT(CHAR(11),GRINDATE,106) GRN_DATE,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,SALESQTY,ISNULL(GRIN_PAGE_NO,0) PAGE_NO, isnull(PerPieceAED, 0) PerPieceAED")
                builder.AppendLine("FROM	TMP_TRADING_INV_GRINS WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "'  ORDER BY SLNO")
            Else
                builder.AppendLine("SELECT TOP 10 PAGE_NO,GRIN_NO,GRN_DATE,SALESQTY,GRINQTY,HOLDQTY,'REMQTY'=AVAILABLEGRNQTY-HOLDQTY,KNOCKOFFQTY,PERPIECEEXCISE,isnull(PerPieceAED, 0) PerPieceAED ")
                builder.AppendLine("FROM")
                builder.AppendLine("(")
                builder.AppendLine("SELECT  ")
                builder.AppendLine("A.PAGE_NO,A.DOC_NO GRIN_NO,")
                builder.AppendLine("CONVERT(CHAR(11),A.GRN_DATE,106) GRN_DATE,")
                builder.AppendLine("'SALESQTY'=Isnull((")
                builder.AppendLine("SELECT TOP 1 SALESQTY FROM TMP_TRADING_INV_GRINS ")
                builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "'")
                builder.AppendLine("),0)")
                builder.AppendLine(",")
                builder.AppendLine("B.ACCEPTED_QUANTITY GRINQTY,")
                builder.AppendLine("B.ACCEPTED_QUANTITY-ISNULL(B.DESPATCH_QTY_TRADING,0) AVAILABLEGRNQTY,'HOLDQTY'=")
                builder.AppendLine("ISNULL((")
                builder.AppendLine("SELECT SUM(KNOCKOFFQTY) KNOCKOFFQTY ")
                builder.AppendLine("FROM   SALESCHALLAN_DTL C,SALES_TRADING_GRIN_DTL D")
                builder.AppendLine("WHERE  C.DOC_NO=D.DOC_NO AND C.UNIT_CODE=D.UNIT_CODE ")
                builder.AppendLine("AND	   A.DOC_NO=D.GRIN_NO ")
                builder.AppendLine("AND    A.DOC_TYPE=D.GRIN_DOC_TYPE")
                builder.AppendLine("AND	   B.ITEM_CODE=D.ITEM_CODE")
                builder.AppendLine("AND    ISNULL(C.BILL_FLAG,0)=0")
                builder.AppendLine("AND    ISNULL(C.CANCEL_FLAG,0)=0")
                builder.AppendLine("AND    C.UNIT_CODE='" + gstrUNITID + "'")
                If ParentFormOperationMode = "EDIT" Then
                    builder.AppendLine("AND  D.DOC_NO<>" + Val(strInvoiceNo).ToString)
                End If
                builder.AppendLine("GROUP  BY D.GRIN_NO,D.GRIN_DOC_TYPE,ITEM_CODE")
                builder.AppendLine("),0),")
                builder.AppendLine("0 KNOCKOFFQTY,")
                builder.AppendLine("ISNULL(PERPIECEEXCISE,0)  PERPIECEEXCISE, ISNULL(PerPieceAED,0)  PerPieceAED")
                builder.AppendLine("FROM GRN_HDR A,GRN_DTL B ")
                builder.AppendLine("WHERE   A.DOC_TYPE=B.DOC_TYPE AND A.DOC_NO=B.DOC_NO")
                builder.AppendLine("AND	 LEN(RTRIM(ISNULL(A.QA_AUTHORIZED_CODE,'')))>0")
                builder.AppendLine("AND	 A.DOC_CATEGORY ='Q' AND A.GRN_CANCELLED=0")
                builder.AppendLine("AND	 A.UNIT_CODE=B.UNIT_CODE")
                builder.AppendLine("AND	 A.UNIT_CODE='" + gstrUNITID + "'")
                builder.AppendLine("AND	 B.ITEM_CODE='" + strInternalPartNo + "'")
                builder.AppendLine(") ABCD WHERE (AVAILABLEGRNQTY-HOLDQTY)>0")
                builder.AppendLine("ORDER BY CAST(GRN_DATE AS DATETIME)")
            End If


            strSql = builder.ToString
            Sqlcmd.CommandText = strSql
            Dr = Sqlcmd.ExecuteReader
            If Dr.HasRows = True Then
                While Dr.Read
                    GridGrin.MaxRows = GridGrin.MaxRows + 1


                    GridGrin.Row = GridGrin.MaxRows
                    GridGrin.Col = GridGrin_ENUM.GRINNO
                    GridGrin.Text = Dr("GRIN_NO")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

                    GridGrin.Col = GridGrin_ENUM.GRINDATE
                    GridGrin.Text = Dr("GRN_DATE")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

                    GridGrin.Col = GridGrin_ENUM.GRINQTY
                    GridGrin.Text = Dr("GRINQTY")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
                    GridGrin.TypeNumberDecPlaces = 2
                    GridGrin.Lock = True

                    GridGrin.Col = GridGrin_ENUM.REMAININGQTY
                    GridGrin.Text = Dr("REMQTY")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
                    GridGrin.TypeNumberDecPlaces = 2
                    GridGrin.Lock = True

                    GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
                    GridGrin.Text = Dr("KNOCKOFFQTY")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
                    GridGrin.TypeNumberMax = Dr("REMQTY") ' REMAINING
                    GridGrin.TypeNumberDecPlaces = 2

                    GridGrin.Col = GridGrin_ENUM.EXCISEPERPIECE
                    GridGrin.Text = Dr("PERPIECEEXCISE")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
                    GridGrin.TypeNumberDecPlaces = 2

                    GridGrin.Col = GridGrin_ENUM.AEDPERPIECE
                    GridGrin.Text = Dr("PerPieceAED")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
                    GridGrin.TypeNumberDecPlaces = 2

                    GridGrin.Col = GridGrin_ENUM.PAGENUMBER
                    GridGrin.Text = Dr("PAGE_NO")
                    GridGrin.CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
                    GridGrin.TypeNumberDecPlaces = 0


                    dblSalesQty = Val(Dr("SALESQTY").ToString)
                End While
                Me.txtSaleQuantity.Text = dblSalesQty
                GridGrin.BlockMode = True
                GridGrin.Row = 1
                GridGrin.Row2 = GridGrin.MaxRows

                GridGrin.Col = 1
                GridGrin.Col2 = GridGrin.MaxCols
                GridGrin.Lock = True
                GridGrin.BlockMode = False

            End If
            If Dr.IsClosed = False Then Dr.Close()
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
            Sqlcmd.Connection.Dispose()
            Sqlcmd.Dispose()
            SQLCon.Dispose()
        End Try
    End Sub
    Private Sub FRMTRADINGINV_A_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            SetBackGroundColorNew(Me, True)
            Me.lblInternalPartNo.Text = strInternalPartNo
            Me.lblCustomerPartNo.Text = strCustPartNo
            Me.lblInternalPartDesc.Text = strInternalPartDesc
            Me.lblCustomerPartDesc.Text = strCustomerPartDesc
            Me.lblCurrentStock.Text = strCurrentStockQty
            PopulateGrins()
            If blnBillFlag = True Then
                Me.txtSaleQuantity.Enabled = False
                Me.btnOK.Enabled = False
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            FRMMKTTRN0076.strGrinAllocationOKCancel = False
            Me.Close()
            Me.Dispose()
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub TotalKnockoff()
        Dim intX As Integer
        Dim dblQuantity As Double
        Dim dblTot As Double
        Dim dblToalExcise As Double
        Dim dblToalAdditional As Double
        GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
        For intX = 1 To GridGrin.MaxRows
            GridGrin.Row = intX
            GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
            dblQuantity = Val(GridGrin.Text)
            dblTot = dblTot + dblQuantity

            GridGrin.Col = GridGrin_ENUM.EXCISEPERPIECE
            dblToalExcise = dblToalExcise + (dblQuantity * Val(GridGrin.Text))

            GridGrin.Col = GridGrin_ENUM.AEDPERPIECE
            dblToalAdditional = dblToalAdditional + (dblQuantity * Val(GridGrin.Text))
        Next
        LblTotalAliocation.Text = dblTot
        LblPendingAliocation.Text = Val(Me.txtSaleQuantity.Text) - Val(LblTotalAliocation.Text)
        Me.lblExciseDuty.Text = dblToalExcise
        Me.lblAdditionalDuty.Text = dblToalAdditional
    End Sub
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim strSql As String
        Dim Sqlcmd As New SqlCommand
        Dim SQLCon As SqlConnection
        Dim intX As Integer
        Dim strPERPIECEEXCISE_1 As String
        Dim strKNOCKOFFQTY_1 As String

        If Val(txtSaleQuantity.Text) = 0 Then
            MsgBox("Sale Quantity Can't Be Zero.", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        If Val(txtSaleQuantity.Text) > Val(Me.lblCurrentStock.Text) Then
            MsgBox("Sale Quantity Can't Be More Than Available Stock.", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        If Val(LblPendingAliocation.Text) > 0 Then
            MsgBox("Sale Quantity Can't Be More Than Stock/Sum of Grin Quantity.", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        For intX = 1 To GridGrin.MaxRows
            GridGrin.Row = intX
            GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
            strKNOCKOFFQTY_1 = GridGrin.Text

            If Val(strKNOCKOFFQTY_1) > 0 Then
                GridGrin.Col = GridGrin_ENUM.EXCISEPERPIECE
                strPERPIECEEXCISE_1 = GridGrin.Text

                If Val(strPERPIECEEXCISE_1) <= 0 Then
                    MessageBox.Show("Bill of Entry against allocated GRIN(s) is pending." + vbCrLf + "Invoice can't be generated.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
            End If
        Next

        SQLCon = SqlConnectionclass.GetConnection()
        Sqlcmd.Connection = SQLCon
        Sqlcmd.CommandType = CommandType.Text
        Dim strGRINNO, strGRINDATE, strGRINQTY, strREMQTY, strKNOCKOFFQTY, strPERPIECEEXCISE, strPageNumber, strPERPIECEAED As String
        Try
            For intX = 1 To GridGrin.MaxRows
                If intX = 1 Then
                    strSql = "DELETE FROM TMP_TRADING_INV_GRINS WHERE IPADDRESS='" + gstrIpaddressWinSck + "' AND UNIT_CODE='" + gstrUNITID + "'"
                    Sqlcmd.CommandText = strSql
                    Sqlcmd.ExecuteNonQuery()
                End If

                GridGrin.Row = intX
                GridGrin.Col = GridGrin_ENUM.GRINNO
                strGRINNO = GridGrin.Text

                GridGrin.Col = GridGrin_ENUM.GRINDATE
                strGRINDATE = GridGrin.Text

                GridGrin.Col = GridGrin_ENUM.GRINQTY
                strGRINQTY = GridGrin.Text

                GridGrin.Col = GridGrin_ENUM.REMAININGQTY
                strREMQTY = GridGrin.Text

                GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
                strKNOCKOFFQTY = GridGrin.Text

                GridGrin.Col = GridGrin_ENUM.EXCISEPERPIECE
                strPERPIECEEXCISE = GridGrin.Text

                GridGrin.Col = GridGrin_ENUM.PAGENUMBER
                strPageNumber = GridGrin.Text

                GridGrin.Col = GridGrin_ENUM.AEDPERPIECE
                strPERPIECEAED = GridGrin.Text


                If Val(strKNOCKOFFQTY) > 0 Then
                    strSql = "INSERT INTO TMP_TRADING_INV_GRINS (SLNO,GRINNO,GRINDATE,ITEM_CODE,SALESQTY,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,IPADDRESS,UNIT_CODE,GRIN_PAGE_NO,PerPieceAED) SELECT " + intX.ToString + "," + strGRINNO + ",'" + strGRINDATE + "','" + lblInternalPartNo.Text + "'," + Me.txtSaleQuantity.Text + "," + strGRINQTY + "," + strREMQTY + "," + strKNOCKOFFQTY + "," + strPERPIECEEXCISE + ",'" + gstrIpaddressWinSck + "','" + gstrUNITID + "'," + Val(strPageNumber).ToString + "," + strPERPIECEAED
                    Sqlcmd.CommandText = strSql
                    Sqlcmd.ExecuteNonQuery()
                End If
            Next
            FRMMKTTRN0076.dblGrinQuantityForSale = Val(Me.LblTotalAliocation.Text)
            FRMMKTTRN0076.lblExciseValue.Text = Me.lblExciseDuty.Text
            FRMMKTTRN0076.lblAEDValue.Text = Me.lblAdditionalDuty.Text
            FRMMKTTRN0076.strGrinAllocationOKCancel = True
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
            Try
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
                Sqlcmd.Connection.Dispose()
                Sqlcmd.Dispose()
                SQLCon.Dispose()
            Catch Ex As Exception
            Finally
                Me.Close()
                Me.Dispose()
            End Try
        End Try
    End Sub
    Public Sub DeleteTmpTable()
        Dim strSql As String
        Dim Sqlcmd As New SqlCommand
        Dim SQLCon As SqlConnection
        SQLCon = SqlConnectionclass.GetConnection()
        Sqlcmd.Connection = SQLCon
        Sqlcmd.CommandType = CommandType.Text
        Try
            strSql = "DELETE FROM TMP_TRADING_INV_GRINS WHERE IPADDRESS='" + gstrIpaddressWinSck + "' AND UNIT_CODE='" + gstrUNITID + "'"
            Sqlcmd.CommandText = strSql
            Sqlcmd.ExecuteNonQuery()
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
            Sqlcmd.Connection.Dispose()
            Sqlcmd.Dispose()
            SQLCon.Dispose()
        End Try
    End Sub

    Private Sub txtSaleQuantity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleQuantity.KeyPress
        Try
            If e.KeyChar = "." Then
                e.Handled = True
            Else
                AllowNumericValueInTextBox(txtSaleQuantity, e)
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub AutoKnockoff()
        Dim intX As Integer
        Dim dblQuantity As Double
        Dim dblRemQuantity As Double
        dblQuantity = Val(Me.txtSaleQuantity.Text)
        GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
        For intX = 1 To GridGrin.MaxRows
            GridGrin.Row = intX
            GridGrin.Col = GridGrin_ENUM.REMAININGQTY
            dblRemQuantity = Val(GridGrin.Text)
            If Val(GridGrin.Text) >= dblQuantity Then
                GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
                GridGrin.Text = dblQuantity.ToString
                dblQuantity = dblQuantity - Val(GridGrin.Text)
            Else
                GridGrin.Col = GridGrin_ENUM.KNOCKOFFQTY
                GridGrin.Text = dblRemQuantity.ToString
                dblQuantity = dblQuantity - dblRemQuantity
            End If
        Next
    End Sub
    Private Sub txtSaleQuantity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSaleQuantity.TextChanged
        Try
            AutoKnockoff()
            TotalKnockoff()
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
End Class