'--------------------------------------------------------------------------------------------------
'COPYRIGHT      :   MIND
'CREATED BY     :   PRAVEEN KUMAR
'CREATED DATE   :   01 JAN 2018
'SCREEN         :   LABEL PRINTING FOR HONDA IN HILEX 
'PURPOSE        :   
'ISSUE ID       :   
'--------------------------------------------------------------------------------------------------
'Modified By         :   Gaurav Kumar
'Modified On         :   27 FEB 2024
'Issue ID            :   INC0109891 — Add Label print functionalty against prod date
'*********************************************************************************************************************

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Collections.Generic
Imports System.IO
Imports System.Text

Public Class FRMMKTTRN0100

    Dim mintFormIndex As Integer
    Private Enum enumMonths
        A = 1
        B
        C
        D
        E
        F
        G
        H
        I
        J
        K
        L
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

    Dim month As enumMonths
    Dim DayOFMonth As Int16
    Dim intCurrentYear As Int16
    Dim strDateCode As String
    Private Sub form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            CtrDisableOnProdDate()
            Call FitToClient(Me, GrpMain, ctlHeader, btnGrpWhiteLabelPrint, 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)

            'Added By priyanka
            rdowith_invoice.Checked = True
            txtItemQty.Enabled = False
            btnInvoiceNoWhiteLabelPrintHelp.Enabled = True
            rbofficeaddress.Checked = True
            rbShippingAddress.Checked = False
            If cboWireLabel.Checked = True Then
                txtLabelQty.Enabled = True
            Else
                txtLabelQty.Enabled = False
            End If


            rdoBarcodePrinter.Checked = True
            cboWireLabel.Checked = True
            cboBoxLabel.Checked = True
            dtpFromDate.Value = GetServerDate()
            dtpToDate.Value = GetServerDate()

            intCurrentYear = GetServerDate.Year
            month = (DateTime.Today).Month
            DayOFMonth = (DateTime.Today).Day
            strDateCode = intCurrentYear.ToString.Substring(2, 2) + month.ToString + DayOFMonth.ToString.PadLeft(2, "0")
            'SetGridsHeader()
            'dtpAuthDate.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Control's Events"

    Private Sub cmdHelpPartNoLabelPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelpPartNoLabelPrint.Click
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Dim strQry As String
        Dim strHelp() As String
        Try

            txtCustPartNo.Text = ""
            txtItemQty.Text = ""
            txtLabelCount.Text = ""
            txtBoxQtyLabelPrint.Text = ""
            If rdowith_invoice.Checked = True Then 'Added By priyanka
                strQry = " SELECT DISTINCT B.ITEM_CODE,C.CUST_DRGNO_FOR_LABELS CUST_DRGNO,Convert(int,B.SALES_QUANTITY) QUANTITY,'' Item_Desc FROM SALESCHALLAN_DTL A INNER JOIN  SALES_DTL B ON A.UNIT_CODE=B.UNIT_CODE AND A.DOC_NO=B.DOC_NO "
                strQry = strQry & " INNER JOIN CUSTITEM_MST C ON A.ACCOUNT_CODE=C.ACCOUNT_CODE AND B.ITEM_CODE=C.ITEM_CODE AND B.UNIT_CODE=C.UNIT_CODE "
                strQry = strQry & " WHERE A.DOC_NO='" & txtInvNoWhiteLabelPrint.Text & "' AND A.UNIT_CODE='" & gstrUNITID & "' AND C.ACTIVE=1 AND A.ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "'  "
            Else
                'INC0109891 Add Drg_Desc column in query
                strQry = " SELECT DISTINCT ITEM_CODE,CUST_DRGNO_FOR_LABELS CUST_DRGNO,'' as QUANTITY,Drg_Desc as Cust_Part_Name FROM CUSTITEM_MST "
                strQry = strQry & " WHERE UNIT_CODE='" & gstrUNITID & "' AND ACTIVE=1 AND ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "'  "
            End If

            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If Not (UBound(strHelp) = -1) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MsgBox("No Record Found !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtCustPartNo.Text = strHelp(1)
                    txtItemQty.Text = strHelp(2)
                    txtLabelCount.Text = strHelp(2)

                    'INC0109891 Start
                    If rdoWithProdDate.Checked Then
                        txtCustPartDesc.Text = strHelp(3)
                    Else
                        txtCustPartDesc.Text = ""
                    End If
                    'INC0109891 End

                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub btnInvoiceNoWhiteLabelPrintHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvoiceNoWhiteLabelPrintHelp.Click
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Dim strQry As String
        Dim strHelp() As String
        Try

            txtInvNoWhiteLabelPrint.Text = ""
            txtCustPartNo.Text = ""
            txtItemQty.Text = ""
            txtLabelCount.Text = ""
            txtBoxQtyLabelPrint.Text = ""

            strQry = " SELECT A.DOC_NO,Invoice_Type FROM SALESCHALLAN_DTL A  WHERE A.UNIT_CODE='" & gstrUNITID & "'  AND A.ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND A.INVOICE_DATE BETWEEN CONVERT(DATE,'" & dtpFromDate.Text & "',103) AND CONVERT(DATE,'" & dtpToDate.Text & "',103) ORDER BY  A.INVOICE_DATE "
            'strQry = " SELECT A.DOC_NO,Invoice_Type FROM SALESCHALLAN_DTL A  WHERE A.UNIT_CODE='" & gstrUNITID & "'  AND A.ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "' ORDER BY  A.INVOICE_DATE "
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If Not (UBound(strHelp) = -1) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MsgBox("No Record Found !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtInvNoWhiteLabelPrint.Text = strHelp(0)
                    'txtCustPartNo.Text = strHelp(2)
                    'txtItemQty.Text = strHelp(3)
                    'txtLabelCount.Text = strHelp(3)
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub cmdHelpCustomerCodeLabelPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelpCustomerCodeLabelPrint.Click
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Dim strQry As String
        Dim strHelp() As String
        Try
            txtCustomerCodeLabelPrint.Text = ""
            txtInvNoWhiteLabelPrint.Text = ""
            txtCustPartNo.Text = ""
            txtItemQty.Text = ""
            txtLabelCount.Text = ""
            txtBoxQtyLabelPrint.Text = ""

            strQry = "SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST where UNIT_CODE ='" & gstrUNITID & "' "
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If Not (UBound(strHelp) = -1) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MsgBox("No Record Found !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtCustomerCodeLabelPrint.Text = strHelp(0)
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

#End Region


    Private Sub txtItemQty_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtItemQty.KeyPress
        Try
            If (Microsoft.VisualBasic.Asc(e.KeyChar) < 48) _
              Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 57) Then
                e.Handled = True
            End If
            If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
                e.Handled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtLabelCount_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLabelCount.KeyPress
        Try
            If (Microsoft.VisualBasic.Asc(e.KeyChar) < 48) _
              Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 57) Then
                e.Handled = True
            End If
            If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
                e.Handled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtLabelQty_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLabelQty.KeyPress
        Try
            If (Microsoft.VisualBasic.Asc(e.KeyChar) < 48) _
              Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 57) Then
                e.Handled = True
            End If
            If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
                e.Handled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtBoxQtyLabelPrint_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBoxQtyLabelPrint.KeyPress
        Try
            If (Microsoft.VisualBasic.Asc(e.KeyChar) < 48) _
              Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 57) Then
                e.Handled = True
            End If
            If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
                e.Handled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Dim strLabelFile As New StringBuilder
    Private Sub printWireLabel()
        Try
            Dim strLabelFormat As String = String.Empty
            Dim strSQL As String = String.Empty
            Dim strLabel As New StringBuilder
            Dim intCnt As Int32
            Dim intTotalLabels As Int32
            Dim strLabelstring As String = String.Empty

            strSQL = "SELECT ISNULL(WIRE_BOX_LABEL,'') FROM BARCODE_CONFIG_MST WHERE UNIT_CODE='" & gstrUNITID & "'"
            strLabelFormat = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
            If String.IsNullOrEmpty(strLabelFormat) Then
                Throw New Exception("Label Format Not Configured !")
                Return
            End If
            strLabelstring = txtCustPartNo.Text.Substring(0, 4) & ">6" & txtCustPartNo.Text.Substring(4, txtCustPartNo.Text.Length - 4)

            strLabel = New StringBuilder("")
            intTotalLabels = Convert.ToInt16(txtLabelCount.Text)
            For intCnt = 1 To intTotalLabels
                strLabel.Append(strLabelFormat)
                strLabel.Replace("V_PARTNO", txtCustPartNo.Text)
                strLabel.Replace("V_DATE", strDateCode)
                strLabel.Replace("V_QTY", txtLabelQty.Text)
                strLabel.Replace("V_BARCODE", strLabelstring)
                strLabel.Replace("V_VENDERCODE", strCUST_VENDOR_CODE)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
            Next

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub printBoxLabel()
        Try
            Dim strLabelFormat As String = String.Empty
            Dim strLabelFormatForCustInfo As String = String.Empty
            Dim strSQL As String = String.Empty
            Dim strLabel As New StringBuilder
            Dim intCnt As Int32
            Dim intTotalLabels As Int32
            Dim intLooseLabel As Int32
            Dim strLabelstring As String = String.Empty

            Dim strCustomerName As String
            Dim arCustomerName(1) As String
            Dim strShipAddress As String
            Dim arShipAddress(5) As String
            Dim strItemDescAndCode As String
            Dim arItemDescAndCode(1) As String
            'Added By priyanka
            Dim strCustomerNameDesc As String
            Dim ADDRESS1 As String
            Dim ADDRESS2 As String
            Dim CITY As String
            Dim district As String
            Dim state As String
            Dim pinno As String
            Dim country As String
            Dim dtcust_detail As New DataTable

            strSQL = "SELECT ISNULL(WIRE_BOX_LABEL,'') FROM BARCODE_CONFIG_MST WHERE UNIT_CODE='" & gstrUNITID & "'"
            strLabelFormat = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
            If String.IsNullOrEmpty(strLabelFormat) Then
                Throw New Exception("Label Format Not Configured !")
                Return
            End If

            strSQL = "SELECT ISNULL(WIRE_BOX_LABEL2,'') FROM BARCODE_CONFIG_MST WHERE UNIT_CODE='" & gstrUNITID & "'"
            strLabelFormatForCustInfo = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
            If String.IsNullOrEmpty(strLabelFormatForCustInfo) Then
                Throw New Exception("Label Format For Customer Information Not Configured !")
                Return
            End If

            intTotalLabels = Math.Truncate(Convert.ToInt16(txtItemQty.Text) / Convert.ToInt16(txtBoxQtyLabelPrint.Text))
            intLooseLabel = Convert.ToInt16(txtItemQty.Text) Mod Convert.ToInt16(txtBoxQtyLabelPrint.Text)
            strLabelstring = txtCustPartNo.Text.Substring(0, 4) & ">6" & txtCustPartNo.Text.Substring(4, txtCustPartNo.Text.Length - 4)
            'Added By priyanka
            If rbShippingAddress.Checked = True Then
                strSQL = "SELECT TOP 1 CUST_NAME,CUST_NAME_DESC,ISNULL(SHIP_ADDRESS1,'') as ADDRESS1,SHIP_ADDRESS2 as ADDRESS2,SHIP_CITY as CITY,Ship_dist as district,SHIP_STATE as state, SHIP_PIN as pinno,SHIP_COUNTRY as country FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            Else
                strSQL = "SELECT TOP 1 CUST_NAME,CUST_NAME_DESC,ISNULL(Office_Address1,'') as ADDRESS1, Office_Address2 as ADDRESS2, Office_City as CITY,Office_dist as district, Office_State as state,Office_Pin as pinno, Office_Country as country FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            End If
            dtcust_detail = SqlConnectionclass.GetDataTable(strSQL)
            strCustomerName = Convert.ToString(dtcust_detail.Rows(0)("CUST_NAME"))
            strCustomerNameDesc = Convert.ToString(dtcust_detail.Rows(0)("CUST_NAME_DESC"))
            ADDRESS1 = Convert.ToString(dtcust_detail.Rows(0)("ADDRESS1"))
            ADDRESS2 = Convert.ToString(dtcust_detail.Rows(0)("ADDRESS2"))
            CITY = Convert.ToString(dtcust_detail.Rows(0)("CITY"))
            district = Convert.ToString(dtcust_detail.Rows(0)("district"))
            state = Convert.ToString(dtcust_detail.Rows(0)("state"))
            pinno = Convert.ToString(dtcust_detail.Rows(0)("pinno"))
            country = Convert.ToString(dtcust_detail.Rows(0)("country"))

            If String.IsNullOrEmpty(strCustomerNameDesc) Then
                arCustomerName = SplitString(strCustomerName, 35)
                strCustomerName = arCustomerName(0)
                strCustomerNameDesc = arCustomerName(1)
            Else
                strCustomerName = strCustomerName.Replace(strCustomerNameDesc, "")
                strCustomerNameDesc = strCustomerNameDesc
            End If

            'strSQL = "SELECT TOP 1 CUST_NAME FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            'strCustomerName = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))

            ''Added By priyanka

            'strSQL = "SELECT TOP 1 CUST_NAME_DESC FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            'strCustomerNameDesc = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))


            'If String.IsNullOrEmpty(strCustomerNameDesc) Then
            '    arCustomerName = SplitString(strCustomerName, 35)
            '    strCustomerName = arCustomerName(0)
            '    strCustomerNameDesc = arCustomerName(1)
            'Else
            '    strCustomerName = strCustomerName.Replace(strCustomerNameDesc, "")
            '    strCustomerNameDesc = strCustomerNameDesc
            'End If

            ''Array.Copy(SplitString(strCustomerName, 30), arCustomerName, arCustomerName.Length)
            ''SplitString(strCustomerName, 30).Copy(arCustomerName)
            'If rbShippingAddress.Checked = True Then
            '    strSQL = "SELECT TOP 1 ISNULL(SHIP_ADDRESS1,'') + ' ' + SHIP_ADDRESS2 + ' ' + SHIP_CITY + ' ' + Office_dist + ' ' + SHIP_STATE + ' ' + SHIP_PIN + ' ' + SHIP_COUNTRY FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            'Else
            '    strSQL = "SELECT TOP 1 ISNULL(Office_Address1,'') + ' ' + Office_Address2 + ' ' + Office_City + ' ' + Office_dist + ' '+ Office_State + ' ' + Office_Pin + ' ' + Office_Country FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            'End If
            'strShipAddress = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
            'arShipAddress = SplitString(strShipAddress, 35)
            ''Array.Copy(SplitString(strShipAddress, 35), arShipAddress, arShipAddress.Length)
            ''SplitString(strShipAddress, 35).Copy(arShipAddress)

            ''strSQL = "	 SELECT Item_Desc + ',' + CUST_DRGNO_FOR_LABELS FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND CUST_DRGNO_FOR_LABELS='" + txtCustPartNo.Text + "' AND ACTIVE=1"

            'Changed By priyanka
            strSQL = "	 SELECT Drg_Desc  FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND CUST_DRGNO_FOR_LABELS='" + txtCustPartNo.Text + "' AND ACTIVE=1"

            strItemDescAndCode = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
            ' arItemDescAndCode = SplitString(strItemDescAndCode, 35) //commented by priyanka

            'Array.Copy(SplitString(strItemDescAndCode, 35), arItemDescAndCode, arItemDescAndCode.Length)
            'SplitString(strItemDescAndCode, 35).Copy(arItemDescAndCode)

            strLabel = New StringBuilder("")
            For intCnt = 1 To intTotalLabels
                strLabel.Append(strLabelFormat)
                strLabel.Replace("V_PARTNO", txtCustPartNo.Text)
                strLabel.Replace("V_DATE", strDateCode)
                strLabel.Replace("V_QTY", txtBoxQtyLabelPrint.Text)
                strLabel.Replace("V_BARCODE", strLabelstring)
                strLabel.Replace("V_VENDERCODE", strCUST_VENDOR_CODE)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
                'For Customer Information Label
                strLabel.Append(strLabelFormatForCustInfo)

                Try
                    strLabel.Replace("V_Custname1", strCustomerName)
                Catch ex As Exception
                    strLabel.Replace("V_Custname1", "")
                End Try
                Try
                    strLabel.Replace("V_Custname2", strCustomerNameDesc)
                Catch ex As Exception
                    strLabel.Replace("V_Custname2", "")
                End Try

                Try
                    strLabel.Replace("V_Ship_Address1", ADDRESS1 + ", " + ADDRESS2)
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address1", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address2", CITY)
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address2", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address3", district + ", " + state + "- " + pinno)
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address3", "")
                End Try

                'Try
                '    strLabel.Replace("V_Ship_Address1", arShipAddress(0))
                'Catch ex As Exception
                '    strLabel.Replace("V_Ship_Address1", "")
                'End Try
                'Try
                '    strLabel.Replace("V_Ship_Address2", arShipAddress(1))
                'Catch ex As Exception
                '    strLabel.Replace("V_Ship_Address2", "")
                'End Try
                'Try
                '    strLabel.Replace("V_Ship_Address3", arShipAddress(2))
                'Catch ex As Exception
                '    strLabel.Replace("V_Ship_Address3", "")
                'End Try
                'Try
                '    strLabel.Replace("V_Ship_Address4", arShipAddress(3))
                'Catch ex As Exception
                '    strLabel.Replace("V_Ship_Address4", "")
                'End Try
                'Try
                '    strLabel.Replace("V_Ship_Address5", arShipAddress(4))
                'Catch ex As Exception
                '    strLabel.Replace("V_Ship_Address5", "")
                'End Try
                'Try
                '    strLabel.Replace("V_Ship_Address6", arShipAddress(5))
                'Catch ex As Exception
                '    strLabel.Replace("V_Ship_Address6", "")
                'End Try
                Try
                    strLabel.Replace("V_Desc1", strItemDescAndCode)
                Catch ex As Exception
                    strLabel.Replace("V_Desc1", "")
                End Try

                'Try
                '    strLabel.Replace("V_Desc1", arItemDescAndCode(0))
                'Catch ex As Exception
                '    strLabel.Replace("V_Desc1", "")
                'End Try
                'Try
                '    strLabel.Replace("V_Desc2", arItemDescAndCode(1))
                'Catch ex As Exception
                '    strLabel.Replace("V_Desc2", "")
                'End Try

                strLabel.Replace("V_Quant", txtBoxQtyLabelPrint.Text)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
            Next
            If intLooseLabel > 0 Then
                strLabel.Append(strLabelFormat)
                strLabel.Replace("V_PARTNO", txtCustPartNo.Text)
                strLabel.Replace("V_DATE", strDateCode)
                strLabel.Replace("V_QTY", intLooseLabel)
                strLabel.Replace("V_BARCODE", strLabelstring)
                strLabel.Replace("V_VENDERCODE", strCUST_VENDOR_CODE)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
                'For Customer Information Label
                strLabel.Append(strLabelFormatForCustInfo)

                Try
                    strLabel.Replace("V_Custname1", arCustomerName(0))
                Catch ex As Exception
                    strLabel.Replace("V_Custname1", "")
                End Try
                Try
                    strLabel.Replace("V_Custname2", arCustomerName(1))
                Catch ex As Exception
                    strLabel.Replace("V_Custname2", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address1", arShipAddress(0))
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address1", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address2", arShipAddress(1))
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address2", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address3", arShipAddress(2))
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address3", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address4", arShipAddress(3))
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address4", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address5", arShipAddress(4))
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address5", "")
                End Try
                Try
                    strLabel.Replace("V_Ship_Address6", arShipAddress(5))
                Catch ex As Exception
                    strLabel.Replace("V_Ship_Address6", "")
                End Try
                Try
                    strLabel.Replace("V_Desc1", arItemDescAndCode(0))
                Catch ex As Exception
                    strLabel.Replace("V_Desc1", "")
                End Try
                Try
                    strLabel.Replace("V_Desc2", arItemDescAndCode(1))
                Catch ex As Exception
                    strLabel.Replace("V_Desc2", "")
                End Try

                strLabel.Replace("V_Quant", intLooseLabel)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
            End If


        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Public Function SplitString(ByVal str As String, ByVal numOfChar As Long) As String()
        Dim sArr() As String
        Dim nCount As Long
        Dim X As Long
        X = Len(str) \ numOfChar
        If X * numOfChar = Len(str) Then ' evently divisible
            ReDim sArr(0 To X)
        Else
            ReDim sArr(0 To X + 1)
        End If
        For X = 1 To Len(str) Step numOfChar
            nCount = nCount + 1
            sArr(nCount - 1) = Mid$(str, X, numOfChar)
        Next
        SplitString = sArr

    End Function
    Dim strCUST_VENDOR_CODE As String
    Private Sub btnGrpWhiteLabelPrint_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles btnGrpWhiteLabelPrint.ButtonClick
        If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            If MsgBox("Do you want to close the screen ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                Me.Close()
                Exit Sub
            Else
                Exit Sub
            End If
        End If
        Call UpdateRegistryDSNProperties(gstrCONNECTIONDSN, gstrCONNECTIONDATABASE, gstrCONNECTIONSERVER)
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)

        If rdoBarcodePrinter.Checked = True Then
            If txtCustomerCodeLabelPrint.Text = "" Then
                MsgBox("Please select customer Code!", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtCustomerCodeLabelPrint.Focus()
                Exit Sub
            End If
            If rdowith_invoice.Checked = True Then 'Added By priyanka
                If txtInvNoWhiteLabelPrint.Text = "" Then
                    MsgBox("Please select invoice No!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtInvNoWhiteLabelPrint.Focus()
                    Exit Sub
                End If
            End If 'Added By priyanka

            If txtCustPartNo.Text = "" Then
                MsgBox("Please select customer Part No!", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtCustPartNo.Focus()
                Exit Sub
            End If

            If txtItemQty.Text = "" Then
                MsgBox("Item Qty Not present!", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtItemQty.Focus()
                Exit Sub
            End If



            strCUST_VENDOR_CODE = SqlConnectionclass.ExecuteScalar("SELECT CUST_VENDOR_CODE FROM CUSTOMER_MST WHERE CUSTOMER_CODE='" & txtCustomerCodeLabelPrint.Text & "' and UNIT_CODE='" & gstrUNITID & "' ")
            strLabelFile.Length = 0

            If cboWireLabel.Checked = True Then
                If txtLabelCount.Text = "" Then
                    strLabelFile.Length = 0
                    MsgBox("Please enter label count!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtLabelCount.Focus()
                    Exit Sub
                End If
                If Convert.ToInt16(txtLabelCount.Text) = 0 Then
                    strLabelFile.Length = 0
                    MsgBox("Please enter valid label count!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtLabelCount.Focus()
                    Exit Sub
                End If
                If txtLabelQty.Text = "" Then
                    strLabelFile.Length = 0
                    MsgBox("Please enter label Qty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtLabelQty.Focus()
                    Exit Sub
                End If
                If Convert.ToInt16(txtLabelQty.Text) = 0 Then
                    strLabelFile.Length = 0
                    MsgBox("Please enter valid label Qty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtLabelCount.Focus()
                    Exit Sub
                End If
                If Convert.ToInt16(txtLabelCount.Text) > Convert.ToInt16(txtItemQty.Text) Then
                    strLabelFile.Length = 0
                    MsgBox("Wire Label Qty can not be greater than total item qty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtLabelCount.Focus()
                    Exit Sub
                End If
            End If
            If cboBoxLabel.Checked = True Then
                If txtBoxQtyLabelPrint.Text = "" Then
                    MsgBox("Please enter Box Qty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    strLabelFile.Length = 0
                    txtBoxQtyLabelPrint.Focus()
                    Exit Sub
                End If
                If Convert.ToInt16(txtBoxQtyLabelPrint.Text) = 0 Then
                    MsgBox("Please enter Valid Box Qty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    strLabelFile.Length = 0
                    txtBoxQtyLabelPrint.Focus()
                    Exit Sub
                End If
                If Convert.ToInt16(txtBoxQtyLabelPrint.Text) > Convert.ToInt16(txtItemQty.Text) Then
                    strLabelFile.Length = 0
                    MsgBox("Box Qty can not be greater than total item qty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtBoxQtyLabelPrint.Focus()
                    Exit Sub
                End If
            End If

            'INC0109891 Start
            If cboWireLabel.Checked = True Then
                If rdoWithProdDate.Checked = True Then
                    printWireLabelByProdDate()
                Else
                    printWireLabel()
                End If

            End If
            If cboBoxLabel.Checked = True Then
                If rdoWithProdDate.Checked = True Then
                    printBoxLabelByProdDate()
                Else
                    printBoxLabel()
                End If
            End If
            'INC0109891 End

            If strLabelFile.Length > 0 Then
                Dim strLabelFilePath As String = gstrUserMyDocPath + "WIRE_BOX_LABELS.TXT"
                Dim strBatchFilePath As String = gstrUserMyDocPath + "WIRE_BOX_LABELS.BAT"
                If File.Exists(strLabelFilePath) Then
                    File.Delete(strLabelFilePath)
                End If
                Using SW As StreamWriter = File.CreateText(strLabelFilePath)
                    SW.Write(strLabelFile)
                    SW.Close()
                End Using
                strLabelFile.Length = 0
                If File.Exists(strBatchFilePath) = False Then
                    Using SW As StreamWriter = File.CreateText(strBatchFilePath)
                        SW.WriteLine("CD\")
                        SW.WriteLine("C:")
                        SW.WriteLine("MODE:LPT1")
                        SW.WriteLine("COPY """ & gstrUserMyDocPath & "WIRE_BOX_LABELS.TXT"" LPT1")
                        SW.Close()
                    End Using
                End If
                Shell("cmd.exe /c """ & gstrUserMyDocPath & "WIRE_BOX_LABELS.BAT""", AppWinStyle.MinimizedNoFocus)
                MsgBox("Barcode Generated Successfully.", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If


    End Sub

    Private Sub rdoBarcodePrinter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoBarcodePrinter.CheckedChanged

    End Sub
    'Added By priyanka
    Private Sub rdowith_invoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdowith_invoice.CheckedChanged
        txtCustomerCodeLabelPrint.Text = ""
        txtInvNoWhiteLabelPrint.Text = ""
        txtItemQty.Enabled = False
        btnInvoiceNoWhiteLabelPrintHelp.Enabled = True
        txtCustPartNo.Text = ""
        txtItemQty.Text = ""
        txtLabelCount.Text = ""
        txtBoxQtyLabelPrint.Text = ""
        rbofficeaddress.Checked = True
        rbShippingAddress.Checked = False
        If cboWireLabel.Checked = True Then
            txtLabelQty.Enabled = True
        Else
            txtLabelQty.Enabled = False
        End If
    End Sub

    Private Sub rdowithout_invoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdowithout_invoice.CheckedChanged
        txtCustomerCodeLabelPrint.Text = ""
        txtInvNoWhiteLabelPrint.Text = ""
        txtItemQty.Enabled = True
        btnInvoiceNoWhiteLabelPrintHelp.Enabled = False
        txtCustPartNo.Text = ""
        txtItemQty.Text = ""
        txtLabelCount.Text = ""
        txtBoxQtyLabelPrint.Text = ""
        rbofficeaddress.Checked = True
        rbShippingAddress.Checked = False
        If cboWireLabel.Checked = True Then
            txtLabelQty.Enabled = True
        Else
            txtLabelQty.Enabled = False
        End If
    End Sub

    Private Sub cboWireLabel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWireLabel.CheckedChanged

        If cboWireLabel.Checked = True Then
            txtLabelQty.Enabled = True
        Else
            txtLabelQty.Enabled = False
        End If
    End Sub

    'INC0109891 Start
    Private Sub rdoWithProdDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoWithProdDate.CheckedChanged
        Try
            txtCustomerCodeLabelPrint.Text = ""
            txtInvNoWhiteLabelPrint.Text = ""
            txtItemQty.Enabled = True
            btnInvoiceNoWhiteLabelPrintHelp.Enabled = False
            txtCustPartNo.Text = ""
            txtItemQty.Text = ""
            txtLabelCount.Text = ""
            txtBoxQtyLabelPrint.Text = ""
            rbofficeaddress.Checked = True
            rbShippingAddress.Checked = False
            If cboWireLabel.Checked = True Then
                txtLabelQty.Enabled = True
            Else
                txtLabelQty.Enabled = False
            End If
            txtCustPartDesc.Text = ""

            If rdoWithProdDate.Checked = True Then
                CtrEnableOnProdDate()
            Else
                CtrDisableOnProdDate()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CtrDisableOnProdDate()
        txtCustPartDesc.Text = ""
        lblPartDesc.Visible = False
        txtCustPartDesc.Visible = False
        lblProdDate.Visible = False
        dtpProdDate.Visible = False
    End Sub
    Private Sub CtrEnableOnProdDate()
        lblPartDesc.Visible = True
        txtCustPartDesc.Visible = True
        lblProdDate.Visible = True
        dtpProdDate.Visible = True
    End Sub
    Private Sub printWireLabelByProdDate()
        Try
            Dim strLabelFormat As String = String.Empty
            Dim strSQL As String = String.Empty
            Dim strLabel As New StringBuilder
            Dim intCnt As Int32
            Dim intTotalLabels As Int32
            Dim Drg_Desc As String

            strSQL = "SELECT ISNULL(WireLable_ProdDate,'') FROM BARCODE_CONFIG_MST WHERE UNIT_CODE='" & gstrUNITID & "'"
            strLabelFormat = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
            If String.IsNullOrEmpty(strLabelFormat) Then
                Throw New Exception("Label Format Not Configured !")
                Return
            End If
            strSQL = "	 SELECT Drg_Desc  FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND CUST_DRGNO_FOR_LABELS='" + txtCustPartNo.Text + "' AND ACTIVE=1"
            Drg_Desc = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))

            strLabel = New StringBuilder("")
            intTotalLabels = Convert.ToInt16(txtLabelCount.Text)

            Dim _lableDate() As String = Convert.ToDateTime(dtpProdDate.Text).ToString("dd-MM-yyyy").Split("-")
            Dim _finalLabelDate As String = _lableDate(0) & ">6-" & _lableDate(1) & "->5" & _lableDate(2)

            For intCnt = 1 To intTotalLabels
                strLabel.Append(strLabelFormat)
                strLabel.Replace("V_PartDesc", Drg_Desc)
                strLabel.Replace("V_PARTNO", txtCustPartNo.Text)
                strLabel.Replace("V_QTY", txtLabelQty.Text)
                strLabel.Replace("V_DATE", Convert.ToDateTime(dtpProdDate.Text).ToString("dd-MM-yyyy"))
                strLabel.Replace("V_BARCODE", _finalLabelDate)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub printBoxLabelByProdDate()
        Try
            Dim strLabelFormat As String = String.Empty
            Dim strSQL As String = String.Empty
            Dim strLabel As New StringBuilder
            Dim intTotalLabels As Int32
            Dim intLooseLabel As Int32
            Dim strLabelstring As String = String.Empty
            Dim intCnt As Int32
            Dim Drg_Desc As String
            strSQL = "SELECT ISNULL(WireBoxLable_ProdDate,'') FROM BARCODE_CONFIG_MST WHERE UNIT_CODE='" & gstrUNITID & "'"
            strLabelFormat = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
            If String.IsNullOrEmpty(strLabelFormat) Then
                Throw New Exception("Label Format Not Configured !")
                Return
            End If

            intTotalLabels = Math.Truncate(Convert.ToInt16(txtItemQty.Text) / Convert.ToInt16(txtBoxQtyLabelPrint.Text))
            intLooseLabel = Convert.ToInt16(txtItemQty.Text) Mod Convert.ToInt16(txtBoxQtyLabelPrint.Text)

            strSQL = "	 SELECT Drg_Desc  FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCodeLabelPrint.Text & "' AND CUST_DRGNO_FOR_LABELS='" + txtCustPartNo.Text + "' AND ACTIVE=1"
            Drg_Desc = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))

            Dim _lableDate() As String = Convert.ToDateTime(dtpProdDate.Text).ToString("dd-MM-yyyy").Split("-")
            Dim _finalLabelDate As String = _lableDate(0) & ">6-" & _lableDate(1) & "->5" & _lableDate(2)

            strLabel = New StringBuilder("")
            For intCnt = 1 To intTotalLabels
                strLabel.Append(strLabelFormat)
                strLabel.Replace("V_PartDesc", Drg_Desc)
                strLabel.Replace("V_PARTNO", txtCustPartNo.Text)
                strLabel.Replace("V_QTY", txtBoxQtyLabelPrint.Text)
                strLabel.Replace("V_DATE", Convert.ToDateTime(dtpProdDate.Text).ToString("dd-MM-yyyy"))
                strLabel.Replace("V_BARCODE", _finalLabelDate)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
            Next
            If intLooseLabel > 0 Then
                strLabel.Append(strLabelFormat)
                strLabel.Replace("V_PartDesc", txtCustPartDesc.Text.Trim())
                strLabel.Replace("V_PARTNO", txtCustPartNo.Text)
                strLabel.Replace("V_QTY", intLooseLabel)
                strLabel.Replace("V_DATE", Convert.ToDateTime(dtpProdDate.Text).ToString("dd-MM-yyyy"))
                strLabel.Replace("V_BARCODE", _finalLabelDate)
                strLabelFile.AppendLine(strLabel.ToString)
                strLabel.Length = 0
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'INC0109891 End
End Class