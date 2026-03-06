Imports System
Imports System.Data.SqlClient
Imports System.IO
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Text
'created by :   Shubhra Verma
'created on :   07 MAY 2011
'Form Name  :   ASN File Generation for RSA

'REVISED BY     :   VINOD SINGH
'REVISION DATE  :   03/06/2011
'REASON         :   CHANGES DONE FOR CHANGED FILE NAMES
'ISSUE ID       :   10101895
'REVISED BY     :   Prashant Dhingra
'REVISION DATE  :   16/06/2011
'REASON         :   Changes done in ASN Test file Generation - DFL Line Added
'ISSUE ID       :   10104491
'REVISED BY     :   Prashant Dhingra
'REVISED DATE   :   27/06/2011
'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
'ISSUE ID       :   10108772 
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   27/01/2012
'ISSUE ID       :   10187030  
'REASON         :   SuppCumQty field is missing in some of the DTL segments
'----------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   19/04/2012
'ISSUE ID       :   1023413   
'REASON         :   Packing Style Description added in DTL Segment of ASN File.
'----------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   SAURAV KUMAR
'REVISED DATE   :   10/05/2012
'ISSUE ID       :   10223098 
'REASON         :   ADDED BAG_QTY AND BOX VALUE in DTL Segment of ASN File.
'----------------------------------------------------------------------------------------------------------------------
'Revised By      : PRASHANT RAJPAL
'Issue ID        : 10237233
'Revision Date   : 15 june 2012
'History         : Changes for ASN  for RSA and Vacuform
'----------------------------------------------------------------------------------------------------------------------
'Revised By      : SHUBHRA VERMA
'Revision Date   : 20 AUG 2012
'History         : MULTI UNIT CHANGES
'----------------------------------------------------------------------------------------------------------------------
'Revised By      : SHUBHRA VERMA
'Revision Date   : 29 JAN 2013
'ISSUE ID        : 10337404 
'DESCRIPTION     : FILE NAME FORMAT IS NOT CORRECT IN ASN FILES  
'----------------------------------------------------------------------------------------------------------------------
'Revised By      : PRASHANT RAJPAL
'Revision Date   : 05-MAR-2013
'ISSUE ID        : 10350863   
'DESCRIPTION     : BIN LABEL NOT PRINT CORRECT-RESOLVED..
'----------------------------------------------------------------------------------------------------------------------
'Revised By      : PRASHANT RAJPAL
'Revision Date   : 03-MAy-2013-07 may 2013
'ISSUE ID        : 10384928   
'DESCRIPTION     : NEW REPORT IS CALLED FROM PREVIEW BUTTON 
'----------------------------------------------------------------------------------------------------------------------
'Revised By      : PRASHANT RAJPAL
'Revision Date   : 15-MAy-2013
'ISSUE ID        : 10389671   
'DESCRIPTION     : FILE ATTRIBUTE IS CHANGED : READ ONLY FILE GENERATED ONLY.
'----------------------------------------------------------------------------------------------------------------------
'Revised By      : Shubhra Verma
'Revision Date   : 27-FEB-2014
'ISSUE ID        : 10549174   
'DESCRIPTION     : ASN Files showing wrong time.
'----------------------------------------------------------------------------------------------------------------------
'Revised By     -  Parveen Kumar
'Revised On     -  31 Oct 2014
'Issue ID       -  10690771.
'Revised History-  ASN Arrival Time.
'****************************************************************************************
'REVISED BY     :  PRASHANT RAJPAL
'REVISED DATE   :  27-jan-2015
'ISSUE ID       :  10713941
'PURPOSE        :  TO INTEGRATE STILLAGE FUNCTIONALITY 
'****************************************************************************************
'REVISED BY     :  PRASHANT RAJPAL
'REVISED DATE   :  08-oct-2015 - 09-oct-2015
'ISSUE ID       :  10912876 
'PURPOSE        :  TO integrate TRIP NO and TRIP selection in ASN FILE
'****************************************************************************************
'Revised By      : Shubhra Verma
'Revision Date   : 08 Mar 2016 
'ISSUE ID        : 10998485
'DESCRIPTION     : Changes in ASN Form for WMART
'****************************************************************************************
'Revised By      : Mayur Kumar
'Revision Date   : 20 Apr 2016
'ISSUE ID        : 101025897
'DESCRIPTION     : Issue in ASN Generation for WMART
'****************************************************************************************
'Revised By      : Milind Mishra
'Revision Date   : 15 Sep 2016
'ISSUE ID        : 101117627
'DESCRIPTION     : Invoices for which ASN have been generated should not get displayed 
'****************************************************************************************
'Revised By      : Anand Yadav
'Revision Date   : 17 Sep 2018
'ISSUE ID        : 101609089
'DESCRIPTION     : ASN Generation for Customer BOSCH and PKC
'****************************************************************************************
'Revised By      : Anand Yadav
'Revision Date   : 05 Dec 2018
'ISSUE ID        : 
'DESCRIPTION     : ASN Generation modification for Customer BOSCH and PKC
'****************************************************************************************
Public Class frmMKTTRN0070
    Private mintFormIndex As Integer
    Dim mblnfilemove As Boolean
    Dim mBkpLocation As String
    Dim mLocalLocation As String
    Dim mShippingDays As Integer
    '10912876
    Dim mblnASnConveyance As Boolean = False
    '10912876


    Private Enum ENUM_ASN
        SEL = 1
        INVOICENO
        INVOICEDATE
    End Enum
    Private Enum ENUM_TRIP
        SEL = 1
        TRIPNO
        DESCRIPTION
    End Enum
    Private Enum ENUM_TRIPTYPE
        SEL = 1
        TRIPTYPE
        DESCRIPTION
    End Enum

    Private Sub cmdCustHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustHelp.Click
        Dim strHelp As String
        Dim strSql As String = ""
        Dim e1 As System.ComponentModel.CancelEventArgs = Nothing

        Try
            strSql = "(select distinct S1.account_code, S1.cust_name " & _
               " from SALESCHALLAN_DTL S1, SALES_DTL S2 " & _
               " where S1.DOC_NO = S2.DOC_NO AND S1.UNIT_CODE= S2.UNIT_CODE AND S1.UNIT_CODE='" & gstrUNITID & "')a"

            strHelp = ShowList(1, 1000, , "S1.ACCOUNT_CODE", "S1.CUST_NAME", "SALESCHALLAN_DTL S1, SALES_DTL S2, CUSTOMER_MST C", " and S1.DOC_NO = S2.DOC_NO " & _
                               " AND S1.UNIT_CODE = C.UNIT_CODE AND S1.Account_code = C.CUSTOMER_CODE " & _
                               " AND S1.CANCEL_FLAG = 0 AND S1.BILL_FLAG = 1 AND S1.UNIT_CODE = S2.UNIT_CODE AND ISNULL(C.PLANT_CODE,'') <> ''  ", "Customer Help", , , , , "S1.UNIT_CODE")

            If strHelp = "-1" Then
                MessageBox.Show("No Customer Code Defined", ResolveResString(100), MessageBoxButtons.OK)
            Else
                txtCustomer.Text = strHelp
                Call txtCustomer_Validating(txtCustomer, e1)
                '10912876
                strSql = "SELECT ASN_CONVEYANCENO FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomer.Text.Trim & "'"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                    mblnASnConveyance = True
                Else
                    mblnASnConveyance = False
                End If
                '10912876
            End If

            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtCustomer_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomer.KeyPress
        Try
            Dim KeyAscii As Short = Asc(e.KeyChar)
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    Call txtCustomer_Validating(txtCustomer, New System.ComponentModel.CancelEventArgs(False))
                Case 39, 34, 96
                    KeyAscii = 0
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomer.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdCustHelp_Click(cmdCustHelp, New System.EventArgs())
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtCustomer_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomer.Validating
        Try
            If txtCustomer.Text.Trim.Length = 0 Then Exit Sub

            Dim strsql As String = ""
            Dim oCmd As SqlCommand
            Dim oRdr As SqlDataReader

            spdInv.MaxRows = 0
            txtTripNo.Text = ""
            lblMessage.Text = ""
            strsql = "SELECT ASN_CONVEYANCENO FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomer.Text.Trim & "'"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                mblnASnConveyance = True
            Else
                mblnASnConveyance = False
            End If
            strsql = "select distinct S1.account_code, S1.cust_name " & _
                     " from SALESCHALLAN_DTL S1, SALES_DTL S2 " & _
                     " where S1.DOC_NO = S2.DOC_NO AND S1.UNIT_CODE= S2.UNIT_CODE AND " & _
                     " S1.UNIT_CODE='" & gStrUnitId & "' AND S1.account_code = '" & txtCustomer.Text & "'"
            'strsql = "SELECT DISTINCT S1.ACCOUNT_CODE, S1.CUST_NAME, S1.DOC_NO, S1.INVOICE_DATE" & _
            '    " FROM SALESCHALLAN_DTL S1, SALES_DTL S2" & _
            '    " WHERE S1.UNIT_CODE = S2.UNIT_CODE AND S1.DOC_NO = S2.DOC_NO" & _
            '    " AND S1.CUST_REF = S2.CUST_REF" & _
            '    " AND S1.CANCEL_FLAG = 0" & _
            '    " AND BILL_FLAG = 1 AND S1.UNIT_CODE = '" & gstrUNITID & "' AND S1.account_code = '" & txtCustomer.Text & "'"

            oCmd = New SqlCommand(strsql, SqlConnectionclass.GetConnection)
            oRdr = oCmd.ExecuteReader(CommandBehavior.CloseConnection)
            If Not oRdr.HasRows Then
                MessageBox.Show("Invalid Customer Code.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            Else
                oRdr.Read()
                lblCustName.Text = oRdr("CUST_NAME").ToString
            End If
            oRdr.Close()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub ADDROW()
        Try
            With spdInv
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = ENUM_ASN.SEL : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .Col = ENUM_ASN.INVOICENO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = ENUM_ASN.INVOICEDATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .set_ColWidth(ENUM_ASN.SEL, 5)
                .set_ColWidth(ENUM_ASN.INVOICENO, 10)
                .set_ColWidth(ENUM_ASN.INVOICEDATE, 20)
            End With
        Catch ex As Exception

        End Try

    End Sub
    Private Sub ADDROW_TRIPNO()
        Try
            With spdtrip
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = ENUM_TRIP.SEL : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .Col = ENUM_TRIP.TRIPNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = ENUM_TRIP.DESCRIPTION : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .set_ColWidth(ENUM_TRIP.SEL, 3)
                .set_ColWidth(ENUM_TRIP.TRIPNO, 6)
                .set_ColWidth(ENUM_TRIP.DESCRIPTION, 20)
            End With
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ADDROW_TRIPTYPE()
        Try
            With spdtriptype
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = ENUM_TRIPTYPE.SEL : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .Col = ENUM_TRIPTYPE.TRIPTYPE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = ENUM_TRIPTYPE.DESCRIPTION : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .set_ColWidth(ENUM_TRIPTYPE.SEL, 3)
                .set_ColWidth(ENUM_TRIPTYPE.TRIPTYPE, 5)
                .set_ColWidth(ENUM_TRIPTYPE.DESCRIPTION, 20)
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Function GenerateASN() As Boolean
        'REVISED BY     :   Prashant Dhingra
        'REVISION DATE  :   16/06/2011
        'REASON         :   Changes done in ASN Test file Generation - DFL Line Added
        'ISSUE ID       :   10104491
        'REVISED BY     :   Prashant Dhingra
        'REVISED DATE   :   27/06/2011
        'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
        'ISSUE ID       :   10108772 
        'REVISED BY     :   Samiksha Tripathi
        'REVISION DATE  :   11/07/2024
        'REASON         :   Changes done for ASN lock out
        'INCIDENTID     :   INC0338430
        Dim rs As IO.StreamReader = Nothing
        Dim readLine As String = Nothing
        Dim strText As String = Nothing
        Dim strTextB As String = Nothing
        Dim strFile As String = Nothing
        Dim i As Integer = 0
        Dim upldFiles As Scripting.File
        Dim SQLCMD As SqlCommand
        Dim SQLRDR As SqlDataReader = Nothing
        Dim STRSQL As String = ""
        Dim spdVal As Object = Nothing
        Dim objFSO As Scripting.FileSystemObject = Nothing
        Dim sqlTran As SqlTransaction
        Dim isTrans As Boolean = False
        Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter
        Dim countInv As Integer
        Dim oFS As New Scripting.FileSystemObject
        Dim rstValidateDB As ClsResultSetDB
        Dim strSQLB As String
        Dim strMessageReferenceNo As String = String.Empty
        Dim strtime As String = String.Empty
        Dim strDeltime As String = String.Empty
        Dim lngCumQty As Long = 0
        Dim dblTotalWeight As Double = 0
        Dim dblItemWeight As Double = 0
        Dim dblTotalConsignmentWeight As Double = 0
        Dim strASNFileString As String = String.Empty   'declared by Vinod
        Dim dtpcummulativedate As String
        '10713941
        Dim blnstillage As Boolean = False
        Dim STRSQLSTILLAGE As String = String.Empty
        Dim blnstillageFunctionality As Boolean = False
        '10713941
        Dim strasnconveyancestring As String
        Dim intcounter As Integer
        Dim spdVal_trip As Object = Nothing
        Dim spdVal_triptype As Object = Nothing
        Dim iloop As Integer

        Try
            If Not ValidateBeforeSave() Then
                gblnCancelUnload = True
                gblnFormAddEdit = True
                Exit Function
            End If
            If (Len(txtTripNo.Text) <> 7) Then
                MsgBox("Please enter Manual Delivery Note Number of length 7.", MsgBoxStyle.Information, ResolveResString(10059))
                Exit Function
            End If

            'SAMIKSHA
            If gstrUNITID = "VF1" Then
                If IsFullSacanEnabled_VF1() Then
                    If ValidateGenerateASN() = False Then
                        gblnCancelUnload = True
                        gblnFormAddEdit = True
                        MsgBox("Could not Generate ASN due to partial BIN V/s ASN Cross Scanning", MsgBoxStyle.Information, ResolveResString(10059))
                        Exit Function
                    End If
                End If
            End If

            If gstrUNITID = "MGS" Then
                If IsFullSacanEnabled_MGS() Then
                    If AllowGenerateASN() = False Then
                        gblnCancelUnload = True
                        gblnFormAddEdit = True
                        MsgBox("Could not Generate ASN due to partial BIN V/s Part & Qty Cross Scanning", MsgBoxStyle.Information, ResolveResString(10059))
                        Exit Function
                    End If
                End If
            End If

            SQLCMD = New SqlCommand
            SQLCMD.Connection = SqlConnectionclass.GetConnection
            sqlTran = SQLCMD.Connection.BeginTransaction
            SQLCMD.Transaction = sqlTran
            isTrans = True

            STRSQL = "Select ASNLoc,BATLoc from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'"
            SQLCMD.CommandText = STRSQL
            SQLRDR = SQLCMD.ExecuteReader

            If SQLRDR.Read Then
                strText = SQLRDR("ASNLoc").ToString
                strTextB = SQLRDR("BATLoc").ToString
            End If
            SQLRDR.Close()

            If Not oFS.FolderExists(strText) Then
                oFS.CreateFolder(strText)
            End If


            'Added by Vinod on 03/06/2011 , Issue id : 10101895
            strASNFileString = GetCustomerFileString()
            'End of addition by Vinod


            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)

            For countInv = 1 To spdInv.MaxRows
                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL
                spdVal = Nothing : spdVal = spdInv.Text
                If spdVal = "" Then spdVal = 0

                If CDbl(spdVal) = 1 Then
                    blnstillage = False
                    spdInv.Row = countInv : spdInv.Col = ENUM_ASN.INVOICENO
                    spdVal = Nothing : spdVal = spdInv.Text
                    'ISSUE ID       :   10223098 


                    STRSQLSTILLAGE = False
                    blnstillageFunctionality = False
                    STRSQL = "select TOP 1 s1.doc_no,ISNULL(C.PKG_STYLECODE,'') PKG_STYLE_CODE,S1.INVOICE_TYPE,S1.SUB_CATEGORY" & _
                        " from saleschallan_dtl s1 inner join sales_dtl s2 on" & _
                        " S1.UNIT_CODE = S2.UNIT_CODE AND s1.doc_no = s2.doc_no " & _
                        " inner join custitem_mst c on " & _
                        " S1.UNIT_CODE = C.UNIT_CODE AND s1.account_code = c.account_code And s2.item_code = c.item_code" & _
                        " inner join customer_mst c1 on " & _
                        " S1.UNIT_CODE = c1.UNIT_CODE AND s1.account_code = c1.Customer_Code " & _
                        " And s2.cust_item_code = c.cust_drgno " & _
                        " inner join item_mst i on I.UNIT_CODE = S2.UNIT_CODE AND i.item_code = s2.item_code" & _
                        " inner join pkg_style_mst p on P.UNIT_CODE = I.UNIT_CODE AND p.pkg_style_c = c.Pkg_stylecode" & _
                        " where S1.UNIT_CODE = '" & gstrUNITID & "' AND s1.cancel_flag = 0 and s1.bill_flag = 1" & _
                        " and s1.doc_no = '" & spdVal & "'" & _
                        " and s1.account_code = '" & txtCustomer.Text & "' and c.Active=1"

                    SQLCMD.CommandText = STRSQL
                    If SQLRDR.IsClosed = False Then
                        SQLRDR.Close()
                    End If
                    SQLRDR = SQLCMD.ExecuteReader()
                    '10713941
                    If SQLRDR.HasRows Then
                        SQLRDR.Read()
                        blnstillageFunctionality = Find_Value("select dbo.UDF_ISTILLAGFUNCTIONALITYENABLED('" & gstrUNITID & "','" & SQLRDR("invoice_type").ToString & "','" & SQLRDR("sub_category").ToString & "')")
                        blnstillage = Find_Value("select dbo.UDF_ISSTILLAGEINVOICE( '" & gstrUNITID & "','" & SQLRDR("invoice_type").ToString & "','" & SQLRDR("sub_category").ToString & "','" & SQLRDR("pkg_style_code").ToString.Trim & "','" & SQLRDR("doc_no").ToString & "')")
                    End If
                    If blnstillageFunctionality = True Then
                        '10713941
                        STRSQL = "select account_code, cust_name,doc_no,invoice_date,Cust_Vendor_Code,Shipping_Duration,Cust_Item_Code, Cust_Item_Desc ," & _
                                " item_code, sales_quantity, pkg_style_des,TO_BOX,binquantity,PKG_STYLE_CODE,INVOICE_TYPE ,unit_code " & _
                                    " FROM VW_STILLAGE_INVOICE_RSA WHERE " & _
                                    " Unit_code='" & gstrUNITID & "' and doc_no ='" & spdVal & "' and account_code = '" & txtCustomer.Text & "' "
                    Else
                        STRSQL = "select s1.account_code, s1.cust_name, s1.doc_no, s1.invoice_date,c1.Cust_Vendor_Code,c1.Shipping_Duration, s2.Cust_Item_Code," & _
                            " s2.Cust_Item_Desc, s2.item_code, s2.sales_quantity, p.pkg_style_des, CASE WHEN C.BINQUANTITY >0 THEN CEILING(SALES_QUANTITY/C.BINQUANTITY )ELSE (S2.TO_BOX-S2.FROM_BOX)+1 END AS 'TO_BOX'  , c.binquantity" & _
                            ",ISNULL(I.PKG_STYLE_C,'') PKG_STYLE_CODE,S1.INVOICE_TYPE,S1.SUB_CATEGORY" & _
                            " from saleschallan_dtl s1 inner join sales_dtl s2 on" & _
                            " S1.UNIT_CODE = S2.UNIT_CODE AND s1.doc_no = s2.doc_no " & _
                            " inner join custitem_mst c on " & _
                            " S1.UNIT_CODE = C.UNIT_CODE AND s1.account_code = c.account_code And s2.item_code = c.item_code" & _
                            " inner join customer_mst c1 on " & _
                            " S1.UNIT_CODE = c1.UNIT_CODE AND s1.account_code = c1.Customer_Code " & _
                            " And s2.cust_item_code = c.cust_drgno " & _
                            " inner join item_mst i on I.UNIT_CODE = S2.UNIT_CODE AND i.item_code = s2.item_code" & _
                            " inner join pkg_style_mst p on P.UNIT_CODE = I.UNIT_CODE AND p.pkg_style_c = i.Pkg_style_c" & _
                            " where S1.UNIT_CODE = '" & gstrUNITID & "' AND s1.cancel_flag = 0 and s1.bill_flag = 1" & _
                            " and s1.doc_no = '" & spdVal & "'" & _
                            " and s1.account_code = '" & txtCustomer.Text & "' and c.Active=1"
                    End If
                    SQLCMD.CommandText = STRSQL
                    If SQLRDR.IsClosed = False Then
                        SQLRDR.Close()
                    End If
                    SQLRDR = SQLCMD.ExecuteReader()

                    i = 0
                    If SQLRDR.HasRows Then
                        SQLRDR.Read()
                        'strFile = strText + txtCustomer.Text + "_" + spdVal
                        strFile = strASNFileString + spdVal + "_" + Format(GetServerDate.Day, "00") + "_" + Format(GetServerDate.Month, "00") + "_" + GetServerDate.Year.ToString

                        oWrite = oFile.CreateText(strText + strFile)

                        i = i + 1
                        'Issue Id - 10104491
                        rstValidateDB = New ClsResultSetDB
                        strSQLB = "Select MessageReferenceNo from ASNMessageReferenceNo WHERE UNIT_CODE = '" & gstrUNITID & "'"
                        Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rstValidateDB.GetNoRows >= 1 Then
                            strMessageReferenceNo = rstValidateDB.GetValue("MessageReferenceNo").ToString
                            mP_Connection.Execute("UPDATE ASNMessageReferenceNo SET MessageReferenceNo = Convert(int, " & strMessageReferenceNo & " ) + 1 WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        rstValidateDB.ResultSetClose()
                        rstValidateDB = Nothing


                        dblTotalConsignmentWeight = 0
                        rstValidateDB = New ClsResultSetDB
                        strSQLB = "Select isnull(sum(Weight*Sales_Quantity),0)/1000 as TotalConsignmentWeight from sales_dtl s inner join item_mst i" & _
                            " on S.UNIT_CODE = I.UNIT_CODE AND s.item_code = i.item_code where S.UNIT_CODE = '" & gstrUNITID & "' AND Doc_no =  '" & SQLRDR("Doc_no") & "'"
                        Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rstValidateDB.GetNoRows >= 1 Then
                            dblTotalConsignmentWeight = Math.Round(Val(rstValidateDB.GetValue("TotalConsignmentWeight")), 4)
                        End If
                        rstValidateDB.ResultSetClose()
                        rstValidateDB = Nothing
                        '10690771
                        'rstValidateDB = New ClsResultSetDB
                        Dim a As Double
                        Dim strtime1 As String = String.Empty
                        '  Dim HHmmss As String = String.Empty
                        a = SQLRDR(5)
                        strtime1 = Date.Parse(DateAdd(DateInterval.Minute, a, GetServerDateTimeNew)).ToString
                        strDeltime = CDate(strtime1).ToString("HHmmss")
                        Dim stras As String
                        stras = DateAdd(DateInterval.Minute, a, GetServerDateTimeNew)

                        '10549174 BEGIN
                        'strSQLB = "Select Replace(Convert(varchar(12),Dateadd(hh,5,dateadd(mi,-30,dateadd(hh,-3,getdate()))),108),':','') as strtime"
                        ' strSQLB = "Select Replace(Convert(varchar(12),strtime1,108),':','') as strtime"
                        '10549174 END
                        'Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        'If rstValidateDB.GetNoRows >= 1 Then
                        '    '  strDeltime = rstValidateDB.GetValue("strtime")


                        'End If
                        'rstValidateDB.ResultSetClose()
                        'rstValidateDB = Nothing

                        '10690771
                        rstValidateDB = New ClsResultSetDB
                        '10549174 Begin
                        'strSQLB = "Select Replace(Convert(varchar(12),dateadd(mi,-30,dateadd(hh,-3,getdate())),108),':','') as strtime"
                        strSQLB = "Select Replace(Convert(varchar(12),getdate(),108),':','') as strtime"
                        '10549174 END
                        Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rstValidateDB.GetNoRows >= 1 Then
                            strtime = rstValidateDB.GetValue("strtime")
                        End If
                        rstValidateDB.ResultSetClose()
                        rstValidateDB = Nothing


                        rstValidateDB = New ClsResultSetDB
                        strSQLB = "Select  Weight from item_mst where UNIT_CODE = '" & gstrUNITID & "' AND item_code =  '" & SQLRDR("item_code") & "'"
                        Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rstValidateDB.GetNoRows >= 1 Then
                            dblItemWeight = rstValidateDB.GetValue("Weight")
                        End If
                        rstValidateDB.ResultSetClose()
                        rstValidateDB = Nothing

                        dblTotalWeight = (dblItemWeight * Math.Round(SQLRDR("sales_quantity"), 4)) / 1000
                        '10690771
                        If (gstrUNITID = "MGS") Then
                            If SQLRDR("account_code") = "C0000014" Then
                                oWrite.WriteLine("DFH,001,GB9BA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(strMessageReferenceNo, "0000000").ToString + ",ASN")
                            Else
                                oWrite.WriteLine("DFH,001,GB9BA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(strMessageReferenceNo, "0000000").ToString + ",ASN")
                            End If
                        ElseIf gstrUNITID = "SA2" Then
                            oWrite.WriteLine("DFH,001,GB9CA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(Val(strMessageReferenceNo), "0000000").ToString + ",ASN")
                        ElseIf gstrUNITID = "VF1" Then
                            oWrite.WriteLine("DFH,001,CGUPA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(Val(strMessageReferenceNo), "0000000").ToString + ",ASN")
                        ElseIf gstrUNITID = "VF2" Then
                            oWrite.WriteLine("DFH,001,CGUPB," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(strMessageReferenceNo, "0000000").ToString + ",ASN")
                        End If
                        '10690771
                        'Issue Id End
                        '10912876
                        If mblnASnConveyance = True Then
                            'If spdtrip.MaxRows > 0 Then

                            '    For iloop = 1 To spdtrip.MaxRows
                            '        spdtrip.Row = iloop : spdtrip.Col = ENUM_TRIP.SEL
                            '        spdVal_trip = Nothing : spdVal_trip = spdtrip.Text
                            '        If spdVal_trip = "" Then spdVal_trip = 0

                            '        If CDbl(spdVal_trip) = 1 Then
                            '            spdtrip.Row = iloop : spdtrip.Col = ENUM_TRIP.TRIPNO
                            '            spdVal_trip = Nothing : spdVal_trip = spdtrip.Text
                            '            Exit For
                            '        End If
                            '    Next
                            'End If

                            'If spdtriptype.MaxRows > 0 Then
                            '    For iloop = 1 To spdtriptype.MaxRows
                            '        spdtriptype.Row = iloop : spdtriptype.Col = ENUM_TRIPTYPE.SEL
                            '        spdVal_triptype = Nothing : spdVal_triptype = spdtriptype.Text
                            '        If spdVal_triptype = "" Then spdVal_triptype = 0

                            '        If CDbl(spdVal_triptype) = 1 Then
                            '            spdtriptype.Row = iloop : spdtriptype.Col = ENUM_TRIPTYPE.TRIPTYPE
                            '            spdVal_triptype = Nothing : spdVal_triptype = spdtriptype.Text
                            '            Exit For
                            '        End If
                            '    Next
                            'End If


                            'strasnconveyancestring = Find_Value("SELECT DBO.UDF_GETASNCONVEYANCE('" & gstrUNITID & "','" & SQLRDR("account_code").ToString & "','" & spdVal_trip & "','" & spdVal_triptype & "')")
                            strasnconveyancestring=txtTripNo.Text

                            '10912876 
                            If (gstrUNITID = "MGS" Or gstrUNITID = "VF1") Then
                                oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," + Convert.ToDateTime(GetServerDate()).Year.ToString + Format(Convert.ToDateTime(GetServerDate()).Month, "00").ToString + Format(Convert.ToDateTime(GetServerDate()).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString + "," + strasnconveyancestring)
                            Else
                                oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," & Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString + "," + strasnconveyancestring)
                            End If
                        Else
                            If (gstrUNITID = "MGS" Or gstrUNITID = "VF1") Then
                                oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," + Convert.ToDateTime(GetServerDate()).Year.ToString + Format(Convert.ToDateTime(GetServerDate()).Month, "00").ToString + Format(Convert.ToDateTime(GetServerDate()).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString)
                            Else
                                oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," & Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString)
                            End If
                        End If



                        lngCumQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_RSA('" & gstrUNITID & "','" & SQLRDR("Cust_Item_Code").ToString & "','" & SQLRDR("account_code").ToString & "'," & SQLRDR("doc_no").ToString & ")")
                        If blnstillage = True Then
                            oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
                            i = i + 1
                            oWrite.WriteLine("DTL," & i & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & SQLRDR("to_box").ToString & ",0," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & ",1")
                        Else
                            oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
                        End If
                        While SQLRDR.Read
                            lngCumQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_RSA('" & gstrUNITID & "','" & SQLRDR("Cust_Item_Code").ToString & "','" & SQLRDR("account_code").ToString & "'," & SQLRDR("doc_no").ToString & ")")

                            rstValidateDB = New ClsResultSetDB
                            strSQLB = "Select  Weight from item_mst where UNIT_CODE = '" & gstrUNITID & "' AND item_code = ' " & SQLRDR("item_code") & "'"
                            Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rstValidateDB.GetNoRows >= 1 Then
                                dblItemWeight = rstValidateDB.GetValue("Weight")
                            End If
                            rstValidateDB.ResultSetClose()
                            rstValidateDB = Nothing

                            dblTotalWeight = dblItemWeight * Math.Round(SQLRDR("sales_quantity"), 4) / 1000
                            i = i + 1
                            '10713941
                            If blnstillage = True Then
                                oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
                                i = i + 1
                                oWrite.WriteLine("DTL," & i & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & SQLRDR("to_box").ToString & ",0," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & ",1")
                            Else
                                oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
                            End If
                            '10713941
                        End While
                        oWrite.Close()
                        SQLRDR.Close()
                    End If
                End If
            Next
            File.SetAttributes(strText + strFile, FileAttributes.ReadOnly)

            If (gstrUserMyDocPath.Substring(gstrUserMyDocPath.Length - 1, 1) = "\") Then
                If File.Exists(gstrUserMyDocPath + strFile) Then
                    File.Delete(gstrUserMyDocPath + strFile)
                End If
                File.Copy(strText + strFile, gstrUserMyDocPath + strFile)
            Else
                If File.Exists(gstrUserMyDocPath + "\" + strFile) Then
                    File.Delete(gstrUserMyDocPath + "\" + strFile)
                End If
                File.Copy(strText + strFile, gstrUserMyDocPath + "\" + strFile)
            End If

            'File.SetAttributes(strText + strFile, FileAttributes.ReadOnly)
            'If File.Exists(gstrUserMyDocPath + strFile) Then
            '    File.Delete(gstrUserMyDocPath + strFile)
            'End If
            'File.Copy(strText + strFile, gstrUserMyDocPath + strFile)

            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)

            MessageBox.Show("ASN Generated SuccessFully", ResolveResString(100), MessageBoxButtons.OK)

            If (gstrUNITID = "MGS" Or gstrUNITID = "VF1") Then
                If File.Exists(strText & "CX_Uploader.BAT") Then
                    Shell("cmd.exe /C """ & strTextB & "CX_Uploader.BAT""", AppWinStyle.MinimizedNoFocus)
                    'Shell(strTextB & "CX_Uploader.bat", AppWinStyle.NormalFocus)
                    MessageBox.Show("ASN Uploaded SuccessFully.", ResolveResString(100), MessageBoxButtons.OK)
                Else
                    MessageBox.Show("ASN can't be Uploaded due to missing Batch File.", ResolveResString(100), MessageBoxButtons.OK)
                End If
            End If
        Catch ex As Exception
            If Not strFile = Nothing Then
                Kill(strFile)
                rs.Dispose()
            End If
            objFSO = Nothing
            If isTrans = True Then
                sqlTran.Rollback()
                isTrans = False
            End If
            Return False
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Function
    'Added by SAMIKSHA
    'Added on 11-07-2024
    'Added for ASN Lock Out VF1
    Private Function IsFullSacanEnabled_VF1() As Boolean
        Dim blnIsFullScan As Boolean = False
        Dim strSql As String = String.Empty

        Try
            strSql = "Select IsNull(IsFullCrossScan,0)IsFullCrossScan from WIPFG_ASN_BIN_CROSS_SCAN_DTL where UNIT_CODE ='" & gstrUNITID & "'"
            strSql = strSql & " and ACCOUNT_CODE='" & txtCustomer.Text.ToString.Trim() & "'"
            blnIsFullScan = SqlConnectionclass.ExecuteScalar(strSql)

            Return blnIsFullScan
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnIsFullScan
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Function
    Private Function IsFullSacanEnabled_MGS() As Boolean
        Dim blnIsFullScan As Boolean = False
        Dim strSql As String = String.Empty

        Try
            strSql = "Select IsNull(IsFullCrossScan,0)IsFullCrossScan from BIN_QTY_PART_CROSS_SCAN_DTL where UNIT_CODE ='" & gstrUNITID & "'"
            strSql = strSql & " and ACCOUNT_CODE='" & txtCustomer.Text.ToString.Trim() & "'"
            blnIsFullScan = SqlConnectionclass.ExecuteScalar(strSql)

            Return blnIsFullScan
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnIsFullScan
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Function

    'Added by SAMIKSHA
    'Added for ASN Generation in MGS
    'ASN will be generated if all the items in the DOC_NO are fully cross scanned

    Private Function AllowGenerateASN() As Boolean
        Dim invoice As String = String.Empty
        Dim count As Long = 0
        Dim intLoopCounter As Int16 = 0
        Dim spdVal As Object = Nothing
        Dim SqlAdp = New SqlDataAdapter
        Dim DSLBLDTL As DataTable
        Dim blnGenASN As Boolean = False
        Dim sqlCmd As New SqlCommand()
        Try
            With Me.spdInv
                For intLoopCounter = 1 To .MaxRows

                    spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.SEL
                    spdVal = Nothing : spdVal = spdInv.Text
                    If spdVal = "" Then spdVal = 0

                    If CDbl(spdVal) = 1 Then

                        spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.INVOICENO
                        spdVal = Nothing : spdVal = spdInv.Text

                        invoice = invoice + spdVal + ";"
                        count = count + 1

                    End If

                Next
            End With
            If count >= 1 Then
                With sqlCmd
                    .CommandText = "USP_VALIDATE_GENERATE_ASN_MGS"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Connection = SqlConnectionclass.GetConnection()
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@INVOICENO", invoice)
                    .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomer.Text.ToString.Trim)
                    .Parameters.Add("@FLAG", SqlDbType.Int).Direction = ParameterDirection.Output
                    .ExecuteNonQuery()
                    If sqlCmd.Parameters("@FLAG").Value.Equals(1) Then
                        blnGenASN = True
                    Else
                        blnGenASN = False
                    End If

                End With
                sqlCmd.Dispose()
                Return blnGenASN


            End If
        Catch ex As Exception
            sqlCmd.Dispose()
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnGenASN
        Finally
            sqlCmd.Dispose()
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        End Try
    End Function
    'Added by SAMIKSHA
    'Added for ASN Generation in VF1
    'ASN will be generated if all the items in the DOC_NO are fully cross scanned

    Private Function ValidateGenerateASN() As Boolean
        Dim invoice As String = String.Empty
        Dim count As Long = 0
        Dim intLoopCounter As Int16 = 0
        Dim spdVal As Object = Nothing
        Dim SqlAdp = New SqlDataAdapter
        Dim DSLBLDTL As DataTable
        Dim blnGenASNAfterCrossScan As Boolean = False
        Dim sqlCmd As New SqlCommand()
        Try

            With Me.spdInv
                For intLoopCounter = 1 To .MaxRows

                    spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.SEL
                    spdVal = Nothing : spdVal = spdInv.Text
                    If spdVal = "" Then spdVal = 0

                    If CDbl(spdVal) = 1 Then

                        spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.INVOICENO
                        spdVal = Nothing : spdVal = spdInv.Text

                        invoice = invoice + spdVal + ";"
                        count = count + 1

                    End If

                Next
            End With
            If count >= 1 Then

                With sqlCmd
                    .CommandText = "USP_VALIDATE_GENERATE_ASN_VF1"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Connection = SqlConnectionclass.GetConnection()
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@INVOICENO", invoice)
                    .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomer.Text.ToString.Trim)
                    .Parameters.Add("@FLAG", SqlDbType.Int).Direction = ParameterDirection.Output
                    .ExecuteNonQuery()
                    If sqlCmd.Parameters("@FLAG").Value.Equals(1) Then
                        blnGenASNAfterCrossScan = True
                    Else
                        blnGenASNAfterCrossScan = False
                    End If

                End With
                sqlCmd.Dispose()
                Return blnGenASNAfterCrossScan
            End If
        Catch ex As Exception
            sqlCmd.Dispose()
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnGenASNAfterCrossScan
        Finally
            sqlCmd.Dispose()
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        End Try

    End Function

    'Private Function GenerateASN() As Boolean
    '    'REVISED BY     :   Prashant Dhingra
    '    'REVISION DATE  :   16/06/2011
    '    'REASON         :   Changes done in ASN Test file Generation - DFL Line Added
    '    'ISSUE ID       :   10104491
    '    'REVISED BY     :   Prashant Dhingra
    '    'REVISED DATE   :   27/06/2011
    '    'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
    '    'ISSUE ID       :   10108772 
    '    Dim rs As IO.StreamReader = Nothing
    '    Dim readLine As String = Nothing
    '    Dim strText As String = Nothing
    '    Dim strTextB As String = Nothing
    '    Dim strFile As String = Nothing
    '    Dim i As Integer = 0
    '    Dim upldFiles As Scripting.File
    '    Dim SQLCMD As SqlCommand
    '    Dim SQLRDR As SqlDataReader = Nothing
    '    Dim STRSQL As String = ""
    '    Dim spdVal As Object = Nothing
    '    Dim objFSO As Scripting.FileSystemObject = Nothing
    '    Dim sqlTran As SqlTransaction
    '    Dim isTrans As Boolean = False
    '    Dim oFile As System.IO.File
    '    Dim oWrite As System.IO.StreamWriter
    '    Dim countInv As Integer
    '    Dim oFS As New Scripting.FileSystemObject
    '    Dim rstValidateDB As ClsResultSetDB
    '    Dim strSQLB As String
    '    Dim strMessageReferenceNo As String = String.Empty
    '    Dim strtime As String = String.Empty
    '    Dim strDeltime As String = String.Empty
    '    Dim lngCumQty As Long = 0
    '    Dim dblTotalWeight As Double = 0
    '    Dim dblItemWeight As Double = 0
    '    Dim dblTotalConsignmentWeight As Double = 0
    '    Dim strASNFileString As String = String.Empty   'declared by Vinod
    '    Dim dtpcummulativedate As String
    '    '10713941
    '    Dim blnstillage As Boolean = False
    '    Dim STRSQLSTILLAGE As String = String.Empty
    '    Dim blnstillageFunctionality As Boolean = False
    '    '10713941
    '    Dim strasnconveyancestring As String
    '    Dim intcounter As Integer
    '    Dim spdVal_trip As Object = Nothing
    '    Dim spdVal_triptype As Object = Nothing
    '    Dim iloop As Integer
    '    Dim strManualTrip As String

    '    Try
    '        If Not ValidateBeforeSave() Then
    '            gblnCancelUnload = True
    '            gblnFormAddEdit = True
    '            Exit Function
    '        End If


    '        SQLCMD = New SqlCommand
    '        SQLCMD.Connection = SqlConnectionclass.GetConnection
    '        sqlTran = SQLCMD.Connection.BeginTransaction
    '        SQLCMD.Transaction = sqlTran
    '        isTrans = True

    '        STRSQL = "Select ASNLoc,BATLoc from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'"
    '        SQLCMD.CommandText = STRSQL
    '        SQLRDR = SQLCMD.ExecuteReader

    '        If SQLRDR.Read Then
    '            strText = SQLRDR("ASNLoc").ToString
    '            strTextB = SQLRDR("BATLoc").ToString
    '        End If
    '        SQLRDR.Close()

    '        If Not oFS.FolderExists(strText) Then
    '            oFS.CreateFolder(strText)
    '        End If


    '        'Added by Vinod on 03/06/2011 , Issue id : 10101895
    '        strASNFileString = GetCustomerFileString()
    '        'End of addition by Vinod


    '        Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)

    '        For countInv = 1 To spdInv.MaxRows
    '            spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL
    '            spdVal = Nothing : spdVal = spdInv.Text
    '            If spdVal = "" Then spdVal = 0

    '            If CDbl(spdVal) = 1 Then
    '                blnstillage = False
    '                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.INVOICENO
    '                spdVal = Nothing : spdVal = spdInv.Text
    '                'ISSUE ID       :   10223098 


    '                STRSQLSTILLAGE = False
    '                blnstillageFunctionality = False
    '                STRSQL = "select TOP 1 s1.doc_no,ISNULL(C.PKG_STYLECODE,'') PKG_STYLE_CODE,S1.INVOICE_TYPE,S1.SUB_CATEGORY" & _
    '                    " from saleschallan_dtl s1 inner join sales_dtl s2 on" & _
    '                    " S1.UNIT_CODE = S2.UNIT_CODE AND s1.doc_no = s2.doc_no " & _
    '                    " inner join custitem_mst c on " & _
    '                    " S1.UNIT_CODE = C.UNIT_CODE AND s1.account_code = c.account_code And s2.item_code = c.item_code" & _
    '                    " inner join customer_mst c1 on " & _
    '                    " S1.UNIT_CODE = c1.UNIT_CODE AND s1.account_code = c1.Customer_Code " & _
    '                    " And s2.cust_item_code = c.cust_drgno " & _
    '                    " inner join item_mst i on I.UNIT_CODE = S2.UNIT_CODE AND i.item_code = s2.item_code" & _
    '                    " inner join pkg_style_mst p on P.UNIT_CODE = I.UNIT_CODE AND p.pkg_style_c = c.Pkg_stylecode" & _
    '                    " where S1.UNIT_CODE = '" & gstrUNITID & "' AND s1.cancel_flag = 0 and s1.bill_flag = 1" & _
    '                    " and s1.doc_no = '" & spdVal & "'" & _
    '                    " and s1.account_code = '" & txtCustomer.Text & "' and c.Active=1"

    '                SQLCMD.CommandText = STRSQL
    '                If SQLRDR.IsClosed = False Then
    '                    SQLRDR.Close()
    '                End If
    '                SQLRDR = SQLCMD.ExecuteReader()
    '                '10713941
    '                If SQLRDR.HasRows Then
    '                    SQLRDR.Read()
    '                    blnstillageFunctionality = Find_Value("select dbo.UDF_ISTILLAGFUNCTIONALITYENABLED('" & gstrUNITID & "','" & SQLRDR("invoice_type").ToString & "','" & SQLRDR("sub_category").ToString & "')")
    '                    blnstillage = Find_Value("select dbo.UDF_ISSTILLAGEINVOICE( '" & gstrUNITID & "','" & SQLRDR("invoice_type").ToString & "','" & SQLRDR("sub_category").ToString & "','" & SQLRDR("pkg_style_code").ToString.Trim & "','" & SQLRDR("doc_no").ToString & "')")
    '                End If
    '                If blnstillageFunctionality = True Then
    '                    '10713941
    '                    STRSQL = "select account_code, cust_name,doc_no,invoice_date,Cust_Vendor_Code,Shipping_Duration,Cust_Item_Code, Cust_Item_Desc ," & _
    '                            " item_code, sales_quantity, pkg_style_des,TO_BOX,binquantity,PKG_STYLE_CODE,INVOICE_TYPE ,unit_code " & _
    '                                " FROM VW_STILLAGE_INVOICE_RSA WHERE " & _
    '                                " Unit_code='" & gstrUNITID & "' and doc_no ='" & spdVal & "' and account_code = '" & txtCustomer.Text & "' "
    '                Else
    '                    STRSQL = "select s1.account_code, s1.cust_name, s1.doc_no, s1.invoice_date,c1.Cust_Vendor_Code,c1.Shipping_Duration, s2.Cust_Item_Code," & _
    '                        " s2.Cust_Item_Desc, s2.item_code, s2.sales_quantity, p.pkg_style_des, CASE WHEN C.BINQUANTITY >0 THEN CEILING(SALES_QUANTITY/C.BINQUANTITY )ELSE (S2.TO_BOX-S2.FROM_BOX)+1 END AS 'TO_BOX'  , c.binquantity" & _
    '                        ",ISNULL(I.PKG_STYLE_C,'') PKG_STYLE_CODE,S1.INVOICE_TYPE,S1.SUB_CATEGORY" & _
    '                        " from saleschallan_dtl s1 inner join sales_dtl s2 on" & _
    '                        " S1.UNIT_CODE = S2.UNIT_CODE AND s1.doc_no = s2.doc_no " & _
    '                        " inner join custitem_mst c on " & _
    '                        " S1.UNIT_CODE = C.UNIT_CODE AND s1.account_code = c.account_code And s2.item_code = c.item_code" & _
    '                        " inner join customer_mst c1 on " & _
    '                        " S1.UNIT_CODE = c1.UNIT_CODE AND s1.account_code = c1.Customer_Code " & _
    '                        " And s2.cust_item_code = c.cust_drgno " & _
    '                        " inner join item_mst i on I.UNIT_CODE = S2.UNIT_CODE AND i.item_code = s2.item_code" & _
    '                        " inner join pkg_style_mst p on P.UNIT_CODE = I.UNIT_CODE AND p.pkg_style_c = i.Pkg_style_c" & _
    '                        " where S1.UNIT_CODE = '" & gstrUNITID & "' AND s1.cancel_flag = 0 and s1.bill_flag = 1" & _
    '                        " and s1.doc_no = '" & spdVal & "'" & _
    '                        " and s1.account_code = '" & txtCustomer.Text & "' and c.Active=1"
    '                End If
    '                SQLCMD.CommandText = STRSQL
    '                If SQLRDR.IsClosed = False Then
    '                    SQLRDR.Close()
    '                End If
    '                SQLRDR = SQLCMD.ExecuteReader()

    '                i = 0
    '                If SQLRDR.HasRows Then
    '                    SQLRDR.Read()
    '                    'strFile = strText + txtCustomer.Text + "_" + spdVal
    '                    strFile = strASNFileString + spdVal + "_" + Format(GetServerDate.Day, "00") + "_" + Format(GetServerDate.Month, "00") + "_" + GetServerDate.Year.ToString

    '                    oWrite = oFile.CreateText(strText + strFile)

    '                    i = i + 1
    '                    'Issue Id - 10104491
    '                    rstValidateDB = New ClsResultSetDB
    '                    strSQLB = "Select MessageReferenceNo from ASNMessageReferenceNo WHERE UNIT_CODE = '" & gstrUNITID & "'"
    '                    Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                    If rstValidateDB.GetNoRows >= 1 Then
    '                        strMessageReferenceNo = rstValidateDB.GetValue("MessageReferenceNo").ToString
    '                        mP_Connection.Execute("UPDATE ASNMessageReferenceNo SET MessageReferenceNo = Convert(int, " & strMessageReferenceNo & " ) + 1 WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    '                    End If
    '                    rstValidateDB.ResultSetClose()
    '                    rstValidateDB = Nothing


    '                    dblTotalConsignmentWeight = 0
    '                    rstValidateDB = New ClsResultSetDB
    '                    strSQLB = "Select isnull(sum(Weight*Sales_Quantity),0)/1000 as TotalConsignmentWeight from sales_dtl s inner join item_mst i" & _
    '                        " on S.UNIT_CODE = I.UNIT_CODE AND s.item_code = i.item_code where S.UNIT_CODE = '" & gstrUNITID & "' AND Doc_no =  '" & SQLRDR("Doc_no") & "'"
    '                    Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                    If rstValidateDB.GetNoRows >= 1 Then
    '                        dblTotalConsignmentWeight = Math.Round(Val(rstValidateDB.GetValue("TotalConsignmentWeight")), 4)
    '                    End If
    '                    rstValidateDB.ResultSetClose()
    '                    rstValidateDB = Nothing
    '                    '10690771
    '                    'rstValidateDB = New ClsResultSetDB
    '                    Dim a As Double
    '                    Dim strtime1 As String = String.Empty
    '                    '  Dim HHmmss As String = String.Empty
    '                    a = SQLRDR(5)
    '                    strtime1 = Date.Parse(DateAdd(DateInterval.Minute, a, GetServerDateTimeNew)).ToString
    '                    strDeltime = CDate(strtime1).ToString("HHmmss")
    '                    Dim stras As String
    '                    stras = DateAdd(DateInterval.Minute, a, GetServerDateTimeNew)

    '                    '10549174 BEGIN
    '                    'strSQLB = "Select Replace(Convert(varchar(12),Dateadd(hh,5,dateadd(mi,-30,dateadd(hh,-3,getdate()))),108),':','') as strtime"
    '                    ' strSQLB = "Select Replace(Convert(varchar(12),strtime1,108),':','') as strtime"
    '                    '10549174 END
    '                    'Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                    'If rstValidateDB.GetNoRows >= 1 Then
    '                    '    '  strDeltime = rstValidateDB.GetValue("strtime")


    '                    'End If
    '                    'rstValidateDB.ResultSetClose()
    '                    'rstValidateDB = Nothing

    '                    '10690771
    '                    rstValidateDB = New ClsResultSetDB
    '                    '10549174 Begin
    '                    'strSQLB = "Select Replace(Convert(varchar(12),dateadd(mi,-30,dateadd(hh,-3,getdate())),108),':','') as strtime"
    '                    strSQLB = "Select Replace(Convert(varchar(12),getdate(),108),':','') as strtime"
    '                    '10549174 END
    '                    Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                    If rstValidateDB.GetNoRows >= 1 Then
    '                        strtime = rstValidateDB.GetValue("strtime")
    '                    End If
    '                    rstValidateDB.ResultSetClose()
    '                    rstValidateDB = Nothing


    '                    rstValidateDB = New ClsResultSetDB
    '                    strSQLB = "Select  Weight from item_mst where UNIT_CODE = '" & gstrUNITID & "' AND item_code =  '" & SQLRDR("item_code") & "'"
    '                    Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                    If rstValidateDB.GetNoRows >= 1 Then
    '                        dblItemWeight = rstValidateDB.GetValue("Weight")
    '                    End If
    '                    rstValidateDB.ResultSetClose()
    '                    rstValidateDB = Nothing

    '                    dblTotalWeight = (dblItemWeight * Math.Round(SQLRDR("sales_quantity"), 4)) / 1000
    '                    '10690771
    '                    If gstrUNITID = "MGS" Then
    '                        If SQLRDR("account_code") = "C0000014" Then
    '                            oWrite.WriteLine("DFH,001,GB9BA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(strMessageReferenceNo, "0000000").ToString + ",ASN")
    '                        Else
    '                            oWrite.WriteLine("DFH,001,GB9BA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(strMessageReferenceNo, "0000000").ToString + ",ASN")
    '                        End If
    '                    ElseIf gstrUNITID = "SA2" Then
    '                        oWrite.WriteLine("DFH,001,GB9CA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(Val(strMessageReferenceNo), "0000000").ToString + ",ASN")
    '                    ElseIf gstrUNITID = "VF1" Then
    '                        oWrite.WriteLine("DFH,001,CGUPA," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(Val(strMessageReferenceNo), "0000000").ToString + ",ASN")
    '                    ElseIf gstrUNITID = "VF2" Then
    '                        oWrite.WriteLine("DFH,001,CGUPB," & SQLRDR("Cust_Vendor_Code").ToString & "," + Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Format(strMessageReferenceNo, "0000000").ToString + ",ASN")
    '                    End If
    '                    '10690771
    '                    'Issue Id End
    '                    '10912876
    '                    If mblnASnConveyance = True Then
    '                        'If spdtrip.MaxRows > 0 Then

    '                        '    For iloop = 1 To spdtrip.MaxRows
    '                        '        spdtrip.Row = iloop : spdtrip.Col = ENUM_TRIP.SEL
    '                        '        spdVal_trip = Nothing : spdVal_trip = spdtrip.Text
    '                        '        If spdVal_trip = "" Then spdVal_trip = 0

    '                        '        If CDbl(spdVal_trip) = 1 Then
    '                        '            spdtrip.Row = iloop : spdtrip.Col = ENUM_TRIP.TRIPNO
    '                        '            spdVal_trip = Nothing : spdVal_trip = spdtrip.Text
    '                        '            Exit For
    '                        '        End If
    '                        '    Next
    '                        'End If

    '                        'If spdtriptype.MaxRows > 0 Then
    '                        '    For iloop = 1 To spdtriptype.MaxRows
    '                        '        spdtriptype.Row = iloop : spdtriptype.Col = ENUM_TRIPTYPE.SEL
    '                        '        spdVal_triptype = Nothing : spdVal_triptype = spdtriptype.Text
    '                        '        If spdVal_triptype = "" Then spdVal_triptype = 0

    '                        '        If CDbl(spdVal_triptype) = 1 Then
    '                        '            spdtriptype.Row = iloop : spdtriptype.Col = ENUM_TRIPTYPE.TRIPTYPE
    '                        '            spdVal_triptype = Nothing : spdVal_triptype = spdtriptype.Text
    '                        '            Exit For
    '                        '        End If
    '                        '    Next
    '                        'End If


    '                        strasnconveyancestring = Find_Value("SELECT DBO.UDF_GETASNCONVEYANCE('" & gstrUNITID & "','" & SQLRDR("account_code").ToString & "','" & spdVal_trip & "','" & spdVal_triptype & "')")
    '                        '10912876 
    '                        If gstrUNITID = "MGS" Then
    '                            oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," + Convert.ToDateTime(GetServerDate()).Year.ToString + Format(Convert.ToDateTime(GetServerDate()).Month, "00").ToString + Format(Convert.ToDateTime(GetServerDate()).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString + "," + strasnconveyancestring + "," + strManualTrip)
    '                        Else
    '                            oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," & Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString + "," + strasnconveyancestring + "," + strManualTrip)
    '                        End If
    '                    Else
    '                        If gstrUNITID = "MGS" Then
    '                            oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," + Convert.ToDateTime(GetServerDate()).Year.ToString + Format(Convert.ToDateTime(GetServerDate()).Month, "00").ToString + Format(Convert.ToDateTime(GetServerDate()).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString)
    '                        Else
    '                            oWrite.WriteLine("HDR," & SQLRDR("cust_name").ToString & "," & gstrUNITDESC & "," & SQLRDR("doc_no").ToString & "," & Format(SQLRDR("invoice_date"), "dd MMM yyyy") & "," & SQLRDR("doc_no").ToString & "," & Convert.ToDateTime(SQLRDR("invoice_date")).Year.ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Month, "00").ToString + Format(Convert.ToDateTime(SQLRDR("invoice_date")).Day, "00").ToString + "," + strtime + "," + Convert.ToDateTime(stras).Year.ToString + Format(Convert.ToDateTime(stras).Month, "00").ToString + Format(Convert.ToDateTime(stras).Day, "00").ToString + "," + strDeltime + "," + dblTotalConsignmentWeight.ToString)
    '                        End If
    '                    End If



    '                    lngCumQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_RSA('" & gstrUNITID & "','" & SQLRDR("Cust_Item_Code").ToString & "','" & SQLRDR("account_code").ToString & "'," & SQLRDR("doc_no").ToString & ")")
    '                    If blnstillage = True Then
    '                        oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
    '                        i = i + 1
    '                        oWrite.WriteLine("DTL," & i & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & SQLRDR("to_box").ToString & ",0," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & ",1")
    '                    Else
    '                        oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
    '                    End If
    '                    While SQLRDR.Read
    '                        lngCumQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_RSA('" & gstrUNITID & "','" & SQLRDR("Cust_Item_Code").ToString & "','" & SQLRDR("account_code").ToString & "'," & SQLRDR("doc_no").ToString & ")")

    '                        rstValidateDB = New ClsResultSetDB
    '                        strSQLB = "Select  Weight from item_mst where UNIT_CODE = '" & gstrUNITID & "' AND item_code = ' " & SQLRDR("item_code") & "'"
    '                        Call rstValidateDB.GetResult(strSQLB, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                        If rstValidateDB.GetNoRows >= 1 Then
    '                            dblItemWeight = rstValidateDB.GetValue("Weight")
    '                        End If
    '                        rstValidateDB.ResultSetClose()
    '                        rstValidateDB = Nothing

    '                        dblTotalWeight = dblItemWeight * Math.Round(SQLRDR("sales_quantity"), 4) / 1000
    '                        i = i + 1
    '                        '10713941
    '                        If blnstillage = True Then
    '                            oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
    '                            i = i + 1
    '                            oWrite.WriteLine("DTL," & i & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & SQLRDR("to_box").ToString & ",0," & Math.Round(dblTotalWeight, 4) & ",STILL," & SQLRDR("to_box").ToString & ",1")
    '                        Else
    '                            oWrite.WriteLine("DTL," & i & "," & SQLRDR("Cust_Item_Code").ToString & "," & SQLRDR("Cust_Item_Desc").ToString & "," & Math.Round(SQLRDR("sales_quantity"), 0).ToString & "," & Math.Round(lngCumQty, 0).ToString & "," & Math.Round(dblTotalWeight, 4) & "," & SQLRDR("pkg_style_des").ToString & "," & SQLRDR("to_box").ToString & "," & SQLRDR("binquantity").ToString & "")
    '                        End If
    '                        '10713941
    '                    End While
    '                    oWrite.Close()
    '                    SQLRDR.Close()
    '                End If
    '                End If
    '        Next
    '        File.SetAttributes(strText + strFile, FileAttributes.ReadOnly)

    '        If (gstrUserMyDocPath.Substring(gstrUserMyDocPath.Length - 1, 1) = "\") Then
    '            If File.Exists(gstrUserMyDocPath + strFile) Then
    '                File.Delete(gstrUserMyDocPath + strFile)
    '            End If
    '            File.Copy(strText + strFile, gstrUserMyDocPath + strFile)
    '        Else
    '            If File.Exists(gstrUserMyDocPath + "\" + strFile) Then
    '                File.Delete(gstrUserMyDocPath + "\" + strFile)
    '            End If
    '            File.Copy(strText + strFile, gstrUserMyDocPath + "\" + strFile)
    '        End If

    '        'File.SetAttributes(strText + strFile, FileAttributes.ReadOnly)
    '        'If File.Exists(gstrUserMyDocPath + strFile) Then
    '        '    File.Delete(gstrUserMyDocPath + strFile)
    '        'End If
    '        'File.Copy(strText + strFile, gstrUserMyDocPath + strFile)

    '        Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)

    '        MessageBox.Show("ASN Generated SuccessFully", ResolveResString(100), MessageBoxButtons.OK)

    '        If gstrUNITID = "MGS" Then
    '            If File.Exists(strText & "CX_Uploader.BAT") Then
    '                Shell("cmd.exe /C """ & strTextB & "CX_Uploader.BAT""", AppWinStyle.MinimizedNoFocus)
    '                'Shell(strTextB & "CX_Uploader.bat", AppWinStyle.NormalFocus)
    '                MessageBox.Show("ASN Uploaded SuccessFully.", ResolveResString(100), MessageBoxButtons.OK)
    '            Else
    '                MessageBox.Show("ASN can't be Uploaded due to missing Batch File.", ResolveResString(100), MessageBoxButtons.OK)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        If Not strFile = Nothing Then
    '            Kill(strFile)
    '            rs.Dispose()
    '        End If
    '        objFSO = Nothing
    '        If isTrans = True Then
    '            sqlTran.Rollback()
    '            isTrans = False
    '        End If
    '        Return False
    '        MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
    '    End Try
    'End Function

    Private Sub frmMKTTRN0069_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            Me.MdiParent = prjMPower.mdifrmMain
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub frmMKTTRN0069_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub frmMKTTRN0069_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, CObj(Panel2), ctlFormHeader1, CObj(Panel1), 250)
            '10998485
            If gstrUnitId = "WMA" Then
                cmdpreview.Visible = False
                spdtrip.Visible = False
                spdtriptype.Visible = False
                lbltripselection.Visible = False
                Label4.Visible = False
                ctlFormHeader1.HeaderString = "ASN Generation"
            Else
                cmdpreview.Visible = True
                'spdtrip.Visible = True
                'spdtriptype.Visible = True
                'lbltripselection.Visible = True
                'Label4.Visible = True
            End If

            txtCustomer.CausesValidation = False

            With spdInv
                .Row = 0
                .Col = ENUM_ASN.SEL : .Text = "Chk"
                .Col = ENUM_ASN.INVOICENO : .Text = "Invoice No."
                .Col = ENUM_ASN.INVOICEDATE : .Text = "Invoice Date"
            End With

            With spdtrip 
                .Row = 0
                .Col = ENUM_TRIP.SEL : .Text = "Chk"
                .Col = ENUM_TRIP.TRIPNO : .Text = "Trip No."
                .Col = ENUM_TRIP.DESCRIPTION : .Text = "Description"
            End With

            With spdtriptype
                .Row = 0
                .Col = ENUM_TRIPTYPE.SEL : .Text = "Chk"
                .Col = ENUM_TRIPTYPE.TRIPTYPE : .Text = "Trip Type"
                .Col = ENUM_TRIPTYPE.DESCRIPTION : .Text = "Description"
            End With

            Call TRIPTYPE()
            Call TRIPNO()

        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub TRIPTYPE()
        Try
            Dim strSql As String = ""
            Dim oCmd As SqlCommand
            Dim oRdr As SqlDataReader

            spdtriptype.MaxRows = 0

            strSql = "SELECT DISTINCT KEY2,Descr FROM LISTS " & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "'" & _
                   " AND Key1='triptype' ORDER BY descr desc"

            oCmd = New SqlCommand
            oCmd.Connection = SqlConnectionclass.GetConnection
            oCmd.CommandText = strSql
            oRdr = oCmd.ExecuteReader

            While oRdr.Read
                Call ADDROW_TRIPTYPE()
                With spdtriptype
                    .Row = .MaxRows : .Col = ENUM_TRIPTYPE.TRIPTYPE : .Text = oRdr("KEY2").ToString
                    .Row = .MaxRows : .Col = ENUM_TRIPTYPE.DESCRIPTION : .Text = oRdr("Descr").ToString
                End With
            End While

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub TRIPNO()
        Try
            Dim strSql As String = ""
            Dim oCmd As SqlCommand
            Dim oRdr As SqlDataReader

            spdtrip.MaxRows = 0

            strSql = "SELECT DISTINCT KEY2,Descr FROM LISTS " & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "'" & _
                   " AND Key1='tripno' ORDER BY key2  "

            oCmd = New SqlCommand
            oCmd.Connection = SqlConnectionclass.GetConnection
            oCmd.CommandText = strSql
            oRdr = oCmd.ExecuteReader

            While oRdr.Read
                Call ADDROW_TRIPNO()
                With spdtrip
                    .Row = .MaxRows : .Col = ENUM_TRIP.TRIPNO : .Text = oRdr("KEY2").ToString
                    .Row = .MaxRows : .Col = ENUM_TRIP.DESCRIPTION : .Text = oRdr("Descr").ToString
                End With
            End While

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Dispose()
    End Sub

    Private Sub cmdGenerateASN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerateASN.Click
        Dim strsql As String
        Dim isBoschCust As Boolean
        Dim isSmrcCust As Boolean
        Dim oDr As SqlDataReader
        Dim strCustCode As String
        Try
            If Not ValidateBeforeSave() Then
                gblnCancelUnload = True
                gblnFormAddEdit = True
                Exit Sub
            End If
            isBoschCust = False
            isSmrcCust = False
            If spdInv.MaxRows > 0 Then

                strsql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST WHERE UNIT_CODE='" & gStrUnitId & "' AND CUST_NAME LIKE '%BOSCH%'"
                oDr = SqlConnectionclass.ExecuteReader(strsql)
                If oDr.HasRows Then
                    While oDr.Read
                        strCustCode = oDr("CUSTOMER_CODE").ToString
                        If strCustCode = txtCustomer.Text Then
                            isBoschCust = True
                        End If
                    End While
                End If
                strsql = "SELECT CUSTOMER_CODE FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUST_NAME LIKE '%SMRC%'"
                oDr = SqlConnectionclass.ExecuteReader(strsql)
                If oDr.HasRows Then
                    While oDr.Read
                        strCustCode = oDr("CUSTOMER_CODE").ToString
                        If strCustCode = txtCustomer.Text Then
                            isSmrcCust = True
                        End If
                    End While
                End If
                If (gstrUNITID = "MS1" Or gstrUNITID = "MP2") And isBoschCust = True Then
                    If BOSCHASNFileGeneration_MTL() = True Then
                        MessageBox.Show("ASN Generated SuccessFully", ResolveResString(100), MessageBoxButtons.OK)
                        PopulateGrid()
                    End If
                ElseIf gstrUNITID = "MS1" And isSmrcCust = True Then
                    If SMRCASNFileGeneration_MTL() = True Then
                        MessageBox.Show("ASN Generated SuccessFully", ResolveResString(100), MessageBoxButtons.OK)
                        PopulateGrid()
                    End If
                ElseIf gstrUNITID = "MS1" Then
                    If PKCASNFileGeneration_MTL() = True Then
                        MessageBox.Show("ASN Generated SuccessFully", ResolveResString(100), MessageBoxButtons.OK)
                        PopulateGrid()
                    End If
                ElseIf gstrUNITID = "WMA" Then
                    If FORDASNFileGeneration_WMART() = True Then
                        MessageBox.Show("ASN Generated SuccessFully", ResolveResString(100), MessageBoxButtons.OK)
                        PopulateGrid()
                    End If
                Else
                    '10912876
                    strsql = "SELECT ASN_CONVEYANCENO ,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomer.Text.Trim & "'"
                    If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                        mblnASnConveyance = True
                    Else
                        mblnASnConveyance = False
                    End If
                    '10912876
                    Call GenerateASN()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub dtpFrom_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpFrom.Validating
        Try
            If dtpFrom.Value > dtpTo.Value Then
                dtpFrom.Value = dtpTo.Value
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub dtpTo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpTo.Validating
        Try
            If dtpTo.Value < dtpFrom.Value Then
                dtpTo.Value = dtpFrom.Value
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdGetInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetInv.Click
        Try
            PopulateGrid()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PopulateGrid()
        Try
            Dim strSql As String = ""
            Dim oCmd As SqlCommand
            Dim oRdr As SqlDataReader
            Call TRIPTYPE()
            Call TRIPNO()

            spdInv.MaxRows = 0

            '101025897 — eMPro - Issue in ASN Generation for WMART
            If gstrUNITID = "WMA" Then
                strSql = "SELECT S1.DOC_NO, S1.INVOICE_DATE " & _
                   " FROM SALESCHALLAN_DTL S1 " & _
                   " WHERE S1.UNIT_CODE='" & gstrUNITID & "' " & _
                   " AND S1.CANCEL_FLAG=0 " & _
                   " AND S1.BILL_FLAG=1 " & _
                   " AND S1.ACCOUNT_CODE='" & txtCustomer.Text & "' " & _
                   " AND INVOICE_DATE BETWEEN '" & Format(dtpFrom.Value, "dd MMM yyyy") & "' AND " & _
                   " '" & Format(dtpTo.Value, "dd MMM yyyy") & "'" & _
                   " AND NOT EXISTS( SELECT * FROM MKT_ASN_INVDTL M WHERE ASN_STATUS=1 AND S1.DOC_NO=M.DOC_NO AND S1.UNIT_CODE=M.UNIT_CODE AND S1.UNIT_CODE='" & gstrUNITID & "') "


            Else
                strSql = "SELECT DISTINCT S1.DOC_NO, S1.INVOICE_DATE" & _
              " FROM SALESCHALLAN_DTL S1, SALES_DTL S2" & _
              " WHERE S1.UNIT_CODE = S2.UNIT_CODE AND S1.UNIT_CODE = '" & gstrUNITID & "'" & _
              " AND S1.DOC_NO = S2.DOC_NO" & _
              " AND S1.CANCEL_FLAG = 0" & _
              " AND BILL_FLAG = 1 AND S1.account_code = '" & txtCustomer.Text & "'" & _
              " AND INVOICE_DATE BETWEEN '" & Format(dtpFrom.Value, "dd MMM yyyy") & "' AND " & _
              " '" & Format(dtpTo.Value, "dd MMM yyyy") & "'"
            End If
            '101025897 — eMPro - Issue in ASN Generation for WMART


            oCmd = New SqlCommand
            oCmd.Connection = SqlConnectionclass.GetConnection
            oCmd.CommandText = strSql
            oRdr = oCmd.ExecuteReader

            While oRdr.Read
                Call ADDROW()
                With spdInv
                    .Row = .MaxRows : .Col = ENUM_ASN.INVOICENO : .Text = oRdr("DOC_NO").ToString
                    .Row = .MaxRows : .Col = ENUM_ASN.INVOICEDATE : .Text = oRdr("INVOICE_DATE").ToString
                End With
            End While
            chkAll.CheckState = 0
        Catch ex As Exception

        End Try
    End Sub
    Private Sub chkAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged
        Try
            Dim i As Integer
            If spdInv.MaxRows > 0 Then
                If chkAll.Checked = True Then
                    For i = 1 To spdInv.MaxRows
                        spdInv.Row = i : spdInv.Col = ENUM_ASN.SEL : spdInv.Text = 1
                    Next
                End If
                If chkAll.Checked = False Then
                    For i = 1 To spdInv.MaxRows
                        spdInv.Row = i : spdInv.Col = ENUM_ASN.SEL : spdInv.Text = 0
                    Next
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Function GetCustomerFileString() As String
        'CREATED BY VINOD ON 03/06/2011
        'ISSUE ID : 10101895
        Dim strSQL As String = String.Empty
        Dim strCustFile As String = ""
        Dim dt As DataTable
        Dim strCompCode As String
        Dim strCustName() As String
        Try
            strSQL = "SELECT COMPANY_CODE FROM COMPANY_MST WHERE UNIT_CODE = '" & gstrUNITID & "'"
            dt = SqlConnectionclass.GetDataTable(strSQL)
            If dt.Rows.Count = 1 Then
                strCompCode = dt.Rows(0).Item("COMPANY_CODE")
            Else
                strCompCode = ""
            End If

            strCustName = lblCustName.Text.Split(" ")
            If strCompCode.Trim = "MSSL" Then ''MSSL
                strCustFile = strCustName(0) + "_MSSL_ASN_"
            ElseIf strCompCode.Trim = "VACF" Then 'VACUFORM
                strCustFile = strCustName(0) + "_Vacuform_ASN_"
            Else
                strCustFile = ""
            End If
            Return strCustFile
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then

            If IsDBNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdpreview.Click
        Try
            If spdInv.MaxRows > 0 Then
                Call GenerateASN_REPORT()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Function GenerateASN_REPORT() As Boolean
        Dim rs As IO.StreamReader = Nothing
        Dim readLine As String = Nothing
        Dim strText As String = Nothing
        Dim strFile As String = Nothing
        Dim i As Integer = 0
        Dim upldFiles As Scripting.File
        Dim SQLCMD As SqlCommand
        Dim SQLRDR As SqlDataReader = Nothing
        Dim STRSQL As String = ""
        Dim spdVal As Object = Nothing
        Dim objFSO As Scripting.FileSystemObject = Nothing
        Dim sqlTran As SqlTransaction
        Dim isTrans As Boolean = False
        Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter
        Dim countInv As Integer
        Dim oFS As New Scripting.FileSystemObject
        Dim rstValidateDB As ClsResultSetDB
        Dim strSQLB As String
        Dim strMessageReferenceNo As String = String.Empty
        Dim strtime As String = String.Empty
        Dim strDeltime As String = String.Empty
        Dim lngCumQty As Long = 0
        Dim dblTotalWeight As Double = 0
        Dim dblItemWeight As Double = 0
        Dim dblTotalConsignmentWeight As Double = 0
        Dim strASNFileString As String = String.Empty   'declared by Vinod
        Dim dtpcummulativedate As String
        Dim strReportName As String
        Try

            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)

            If Not ValidateBeforeSave() Then
                gblnCancelUnload = True
                gblnFormAddEdit = True
                Exit Function
            End If

            For countInv = 1 To spdInv.MaxRows
                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL
                spdVal = Nothing : spdVal = spdInv.Text
                If spdVal = "" Then spdVal = 0

                If CDbl(spdVal) = 1 Then
                    spdInv.Row = countInv : spdInv.Col = ENUM_ASN.INVOICENO
                    spdVal = Nothing : spdVal = spdInv.Text

                    Dim objRpt As ReportDocument
                    Dim strRepPath As String
                    Dim frmReportViewer As New eMProCrystalReportViewer
                    objRpt = frmReportViewer.GetReportDocument()
                    frmReportViewer.ShowPrintButton = True
                    frmReportViewer.ShowTextSearchButton = True
                    frmReportViewer.ShowZoomButton = True
                    frmReportViewer.ReportHeader = "FORD ASN FILE "

                    With objRpt
                        strReportName = "\Reports\FordRSAGeneration" & GetPlantName() & ".rpt"
                        If Not CheckFile(strReportName) Then
                            strReportName = "\Reports\FordRSAGeneration.rpt"
                        End If


                        frmReportViewer.ExcelExportOption_Column_width = -1
                        '-------------------------------------------------'
                        .Load(My.Application.Info.DirectoryPath & strReportName)

                        .RecordSelectionFormula = "{VW_FORDASNGENERATION.unit_code} = '" & gstrUNITID & "'and  {VW_FORDASNGENERATION.doc_no} = " & spdVal & ""
                        frmReportViewer.Zoom = 120
                        frmReportViewer.Show()
                    End With
                    '<<<<CR11 Code Ends>>>>
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

                End If
            Next

            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)


        Catch ex As Exception
            If Not strFile = Nothing Then
                Kill(strFile)
                rs.Dispose()
            End If
            objFSO = Nothing
            If isTrans = True Then
                sqlTran.Rollback()
                isTrans = False
            End If
            Return False
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Function
    Private Function ValidateBeforeSave() As Boolean
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        Dim countInv As Integer
        Dim spdVal As Object = Nothing
        Dim intcounter As Integer
        ValidateBeforeSave = True
        intcounter = 0

        lstrControls = ResolveResString(10059)
        If (Len(txtCustomer.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Customer code ."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = txtCustomer
            End If
            ValidateBeforeSave = False
        End If
        If spdInv.MaxRows > 0 Then
            For countInv = 1 To spdInv.MaxRows
                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL
                spdVal = Nothing : spdVal = spdInv.Text
                If spdVal = "" Then spdVal = 0

                If CDbl(spdVal) = 1 Then
                    intcounter = intcounter + 1
                    If intcounter > 1 Then
                        ValidateBeforeSave = False
                        lstrControls = lstrControls & vbCrLf & ". MORE THAN 1 INVOICES CAN'T SELECT ."
                        Exit For
                    End If
                End If
            Next
        End If
        If intcounter <= 0 Then
            ValidateBeforeSave = False
            lstrControls = lstrControls & vbCrLf & ". SELECT INVOICE NO ."
        End If

        '10912876
        'If mblnASnConveyance = True Then
        '    If spdtrip.MaxRows > 0 Then
        '        intcounter = 0
        '        For countInv = 1 To spdtrip.MaxRows
        '            spdtrip.Row = countInv : spdtrip.Col = ENUM_TRIP.SEL
        '            spdVal = Nothing : spdVal = spdtrip.Text
        '            If spdVal = "" Then spdVal = 0

        '            If CDbl(spdVal) = 1 Then
        '                intcounter = intcounter + 1
        '                If intcounter > 1 Then
        '                    ValidateBeforeSave = False
        '                    lstrControls = lstrControls & vbCrLf & ". MORE THAN 1 TRIP CAN'T SELECT ."
        '                    Exit For
        '                End If
        '            End If
        '        Next
        '    End If
        '    If intcounter <= 0 Then
        '        ValidateBeforeSave = False
        '        lstrControls = lstrControls & vbCrLf & ". SELECT TRIP NO ."
        '    End If
        'End If
        '10912876
        'If mblnASnConveyance = True Then
        '    If spdtriptype.MaxRows > 0 Then
        '        intcounter = 0
        '        For countInv = 1 To spdtriptype.MaxRows
        '            spdtriptype.Row = countInv : spdtriptype.Col = ENUM_TRIP.SEL
        '            spdVal = Nothing : spdVal = spdtriptype.Text
        '            If spdVal = "" Then spdVal = 0
        '            If CDbl(spdVal) = 1 Then
        '                intcounter = intcounter + 1
        '                If intcounter > 1 Then
        '                    ValidateBeforeSave = False
        '                    lstrControls = lstrControls & vbCrLf & ". MORE THAN 1 TRIP TYPE CAN'T SELECT ."
        '                    Exit For
        '                End If
        '            End If
        '        Next
        '    End If
        '    If intcounter <= 0 Then
        '        ValidateBeforeSave = False
        '        lstrControls = lstrControls & vbCrLf & ". SELECT TRIP TYPE ."
        '    End If
        'End If

        'ASN CONVEYANCE


        If Not ValidateBeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
        End If

        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Function
    End Function

    Private Sub spdtrip_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles spdtrip.ClickEvent
        Dim irow As Integer
        Dim icnt As Integer

        On Error GoTo ErrHandler

        If spdtrip.MaxRows <= 0 Then Exit Sub

        If spdtrip.ActiveCol = 1 Then
            irow = spdtrip.ActiveRow

            For icnt = 1 To spdtrip.MaxRows
                spdtrip.Row = icnt
                spdtrip.Col = 1
                If spdtrip.Row <> irow Then
                    spdtrip.Value = 0
                End If
            Next
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub spdtriptype_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles spdtriptype.ClickEvent
        Dim irow As Integer
        Dim icnt As Integer

        On Error GoTo ErrHandler
        If spdtriptype.MaxRows <= 0 Then Exit Sub

        If spdtriptype.ActiveCol = 1 Then
            irow = spdtriptype.ActiveRow

            For icnt = 1 To spdtriptype.MaxRows
                spdtriptype.Row = icnt
                spdtriptype.Col = 1
                If spdtriptype.Row <> irow Then
                    spdtriptype.Value = 0
                End If
            Next

        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Function FORDASNFileGeneration_WMART() As String
        '10998485

        Dim intCount As Short
        Dim STRPLANTCODE As String
        Dim strFileName As String
        Dim intLineNo As Short
        Dim inttotalNoofItems As Integer
        Dim strSql As String
        Dim strRecord As String
        Dim rsgetData As New ClsResultSetDB
        Dim rsEx_Cess As ClsResultSetDB
        Dim rsSalesDtl As ClsResultSetDB
        Dim strInvoice_item As String
        Dim dblTotalExciseAmt As Double
        Dim dblSalesQty As Double
        Dim strcontainerdespQty As String
        Dim dblcummulativeQty As Double
        Dim dblContainerQty As Double
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim FSO As New Scripting.FileSystemObject
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim objDocno As Object = Nothing
        Dim countInv As Integer = 0
        Dim objSel As Object = Nothing
        Dim strupdateASNdtl As String = String.Empty
        Dim strupdateASNCumFig As String = String.Empty

        Try
            For countInv = 1 To spdInv.MaxRows
                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL : objSel = Nothing : objSel = spdInv.Text

                If objSel = "" Then objSel = 0

                If CDbl(objSel) = 1 Then

                    spdInv.Row = countInv : spdInv.Col = ENUM_ASN.INVOICENO
                    objDocno = Nothing : objDocno = spdInv.Text

                    inttotalNoofItems = Find_Value("SELECT DISTINCT COUNT(CUST_ITEM_CODE) CUST_ITEM_CODE FROM SALES_DTL (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' and DOC_NO='" & objDocno & "'")
                    STRPLANTCODE = Trim(Find_Value("SELECT ISNULL(PLANT_CODE ,'') FROM customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' and customer_code='" & txtCustomer.Text & "'"))

                    If Len(STRPLANTCODE) > 0 Then
                        strSql = ""
                        strSql = "select * from dbo.FN_GETASNDETAIL_WMART(" & objDocno & ",'" & gstrUnitId & "')"
                        rsgetData = New ClsResultSetDB
                        rsgetData.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        rsgetData.MoveFirst()
                        strRecord = ""
                        Do While Not rsgetData.EOFRecord
                            dblcummulativeQty = 0
                            dblSalesQty = 0
                            dblContainerQty = 0

                            dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY('" & gstrUnitId & "','" & STRPLANTCODE & "','" & rsgetData.GetValue("cust_item_code").ToString() & "'," & objDocno & ")")
                            dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
                            dblcummulativeQty = dblcummulativeQty + dblSalesQty
                            strRecord = strRecord & rsgetData.GetValue("cust_vendor_code").ToString 'SUPPLIER CODE
                            strRecord = strRecord & "," & rsgetData.GetValue("CUSTOMERCODE").ToString 'CUSTOMER CODE
                            strRecord = strRecord & "," & objDocno
                            strRecord = strRecord & "," & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "yyyymmddhhmmss")
                            strRecord = strRecord & "," & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "yyyymmddhhmmss")
                            strRecord = strRecord & "," & rsgetData.GetValue("GROSS_WEIGHT").ToString
                            strRecord = strRecord & "," & rsgetData.GetValue("NET_WEIGHT").ToString
                            strRecord = strRecord & "," & rsgetData.GetValue("CONS_MEASURE_CODE").ToString 'Measurement Unit
                            strRecord = strRecord & "," & rsgetData.GetValue("CUST_PLANTCODE").ToString
                            strRecord = strRecord & "," & rsgetData.GetValue("SHIP_PLANTCODE").ToString
                            strRecord = strRecord & "," & rsgetData.GetValue("TRANSPORT_TYPE").ToString 'TRANSPORT TYPE
                            strRecord = strRecord & "," & rsgetData.GetValue("EQUIPEMENTQUALIFIER").ToString 'TRANSPORT TYPE
                            strRecord = strRecord & "," & rsgetData.GetValue("CONVEYANCENO").ToString '13
                            strRecord = strRecord & "," & rsgetData.GetValue("Container").ToString '14
                            strRecord = strRecord & "," & rsgetData.GetValue("CONTAINER_DESP_QTY").ToString '15
                            strRecord = strRecord & "," & rsgetData.GetValue("NO_OF_PKGS").ToString '16
                            strRecord = strRecord & "," & rsgetData.GetValue("QTY_PER_BAG").ToString '17
                            strRecord = strRecord & "," & rsgetData.GetValue("NO_OF_PKGS").ToString '18
                            strRecord = strRecord & "," & rsgetData.GetValue("CUST_ITEM_CODE").ToString 'Part Number 19
                            strRecord = strRecord & "," & rsgetData.GetValue("sales_quantity").ToString 'Despatch Qty 20
                            strRecord = strRecord & "," & dblcummulativeQty
                            strRecord = strRecord & "," & rsgetData.GetValue("MEASURE_UNIT").ToString 'Measurement Unit 22
                            strRecord = strRecord & "," & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "yyyymmddhhmmss") '23
                            strRecord = strRecord & "," & objDocno '24
                            strRecord = strRecord & "," & "" '25
                            strRecord = strRecord & "," & rsgetData.GetValue("CUSTOMERCODE").ToString 'CUSTOMER CODE '26
                            strRecord = strRecord & "," & objDocno '27
                            strRecord = strRecord & "," & "" '28
                            strRecord = strRecord & "," & rsgetData.GetValue("Total_QTY").ToString 'Despatch Qty
                            strRecord = strRecord & "," & inttotalNoofItems & vbCrLf '30
                            intLineNo = intLineNo + 1

                            strupdateASNdtl = "IF EXISTS (SELECT ASN_STATUS FROM MKT_ASN_INVDTL " & _
                                                 " WHERE(DOC_NO = " & objDocno & " " & _
                                                 " AND UNIT_CODE='" & gstrUnitId & "') ) " & _
                                                 " BEGIN " & _
                                                 " UPDATE MKT_ASN_INVDTL SET ASN_STATUS = 1 " & _
                                                 " WHERE UNIT_CODE='" & gstrUnitId & "' AND DOC_NO = '" & objDocno & "' " & _
                                                 " End " & _
                                                 " ELSE BEGIN " & _
                                                 " INSERT INTO MKT_ASN_INVDTL (DOC_NO, CUST_PLANTCODE, ASN_STATUS, CUST_PART_CODE, CUMMULATIVE_QTY, UNIT_CODE)" & _
                                                 " VALUES ( " & objDocno & ", " & _
                                                 " '" & STRPLANTCODE & "'," & _
                                                 "  1, '" & rsgetData.GetValue("CUST_ITEM_CODE").ToString().Trim() & "', " & _
                                                 " " & dblcummulativeQty & ", '" & gstrUnitId & "') " & _
                                                 " End "

                            SqlConnectionclass.ExecuteNonQuery(strupdateASNdtl)

                            strupdateASNCumFig = "IF NOT EXISTS(SELECT CUST_PART_CODE FROM MKT_ASN_CUMFIG " & _
                                " WHERE CUST_PLANTCODE='" & STRPLANTCODE & "' AND UNIT_CODE='" & gstrUnitId & "')" & _
                                " BEGIN    INSERT INTO MKT_ASN_CUMFIG(CUST_PART_CODE,CUST_PLANTCODE,CUMMULATIVE_QTY,UNIT_CODE) " & _
                                " VALUES('" & rsgetData.GetValue("CUST_ITEM_CODE").ToString().Trim() & "' ,'" & STRPLANTCODE & "'," & dblcummulativeQty & ", '" & gstrUnitId & "')" & _
                                " End ELSE " & _
                                " BEGIN " & _
                                " UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY = " & dblcummulativeQty & " " & _
                                " WHERE CUST_PART_CODE = '" & rsgetData.GetValue("CUST_ITEM_CODE").ToString().Trim() & "' " & _
                                " AND CUST_PLANTCODE='" & STRPLANTCODE & "' AND UNIT_CODE='" & gstrUnitId & "'" & _
                                " End"
                            SqlConnectionclass.ExecuteNonQuery(strupdateASNCumFig)

                            'strupdateASNdtl = Trim(strupdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE DOC_NO=" & objDocno & " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_ITEM_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "' AND UNIT_CODE='" & gstrUnitId & "'" & vbCrLf
                            'strupdateASNCumFig = Trim(strupdateASNCumFig) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE CUST_PART_CODE='" & rsgetData.GetValue("CUST_ITEM_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "' AND UNIT_CODE='" & gstrUNITID & "'" & vbCrLf
                            rsgetData.MoveNext()
                        Loop

                        gstrASNPath = gstrUserMyDocPath
                        gstrASNPathForEDI = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUnitId, "FilepathforEDI")
                        If Directory.Exists(gstrASNPath) = False Then
                            Directory.CreateDirectory(gstrASNPath)
                        End If
                        If Directory.Exists(gstrASNPathForEDI) = False Then
                            Directory.CreateDirectory(gstrASNPathForEDI)
                        End If
                        strASNFilepath = gstrASNPath & "\" & objDocno.ToString() & ".CSV"
                        strASNFilepathforEDI = gstrASNPathForEDI & "\" & objDocno.ToString() & ".CSV"
                        fs = File.Create(strASNFilepath)
                        sw = New StreamWriter(fs)
                        'sw.WriteLine(strRecord)
                        sw.Write(strRecord.TrimEnd(New Char() {vbCr, vbLf}))

                        sw.Close()
                        fs.Close()

                        If File.Exists(strASNFilepathforEDI) = False Then
                            File.Copy(strASNFilepath, strASNFilepathforEDI)
                        End If
                        rsgetData.ResultSetClose()
                        rsgetData = Nothing
                        FORDASNFileGeneration_WMART = True
                        ' Exit Function
                    Else
                        MessageBox.Show("Unable To Get Plant Code For The Customer: " & txtCustomer.Text & " While Generating ASN File." & vbCrLf & _
                                            "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        FORDASNFileGeneration_WMART = False

                        Exit Function

                    End If

                End If

            Next
            Exit Function


        Catch ex As Exception
            RaiseException(ex)
        Finally
            FileClose(1)
        End Try

    End Function
    Private Function BOSCHASNFileGeneration_MTL() As String
        Dim intCount As Short
        Dim STRPLANTCODE As String
        Dim strSql As String
        Dim strRecord As String
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim FSO As New Scripting.FileSystemObject
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim objDocno As Object = Nothing
        Dim countInv As Integer = 0
        Dim objSel As Object = Nothing
        Dim oRDR_Hdr As SqlDataReader
        Dim oRDR_Dtl As SqlDataReader
        Dim STRCUSTVENDCODE As String

        Try
            BOSCHASNFileGeneration_MTL = False
            For countInv = 1 To spdInv.MaxRows
                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL : objSel = Nothing : objSel = spdInv.Text
                If objSel = "" Then objSel = 0
                If CDbl(objSel) = 1 Then
                    intCount += 1
                    spdInv.Row = countInv : spdInv.Col = ENUM_ASN.INVOICENO
                    objDocno = Nothing : objDocno = spdInv.Text
                    STRPLANTCODE = Trim(Find_Value("SELECT ISNULL(PLANT_CODE ,'') FROM customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' and customer_code='" & txtCustomer.Text & "'"))
                    STRCUSTVENDCODE = Trim(Find_Value("SELECT ISNULL(cust_vendor_code ,'') FROM customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' and customer_code='" & txtCustomer.Text & "'"))
                    If Len(STRPLANTCODE) > 0 Then
                        strSql = ""
                        If gstrUNITID.ToUpper = "MS1" Or gstrUNITID.ToUpper = "MS2" Then
                            strSql = " select 'HDR'+','+'0097359171'+','+'BOSCH EDI-TEAM'+','+'53F0'+','+ CAST(h.Doc_No AS VARCHAR(10)) + ','" & _
                        " +CAST(DBO.FN_DATEFORMAT_YEARMONTH(INVOICE_DATE) AS VARCHAR(6)) " & _
                        " + CASE WHEN LEN(DAY(INVOICE_DATE)) = 1 THEN '0' + CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) ELSE  CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) END" & _
                        " + replace(convert(varchar(8),h.ent_dt,108),':', '') +', '+CAST(h.Doc_No AS VARCHAR(10))+','" & _
                        " +'53F0'+','+'CLJ01'+','+'0097359171,53F0'+','+'M'+','+CAST(h.Doc_No AS VARCHAR(10))+','+d.cust_item_code+','+" & _
                        " d.item_code+','+'03'+','+ CASE WHEN LEN(DAY(INVOICE_DATE)) = 1 THEN '0' + CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) ELSE  CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) END + " & _
                        " CASE WHEN LEN(Month(INVOICE_DATE)) = 1 THEN '0' + CAST(Month(INVOICE_DATE) AS VARCHAR(2)) ELSE  CAST(Month(INVOICE_DATE) AS VARCHAR(2)) END +" & _
                        " CAST(year(INVOICE_DATE) AS VARCHAR(4))+','+ cast(d.sales_quantity as varchar(20))+',PCE,'+CAST(d.cust_ref AS VARCHAR(25))+','+isnull(h.final_destination,'')+" & _
                        " ','+'0097359171'+','+'5101'  from saleschallan_dtl h inner join sales_dtl d on" & _
                        " h.unit_code = d.unit_code And h.doc_no = d.doc_no where h.UNIT_CODE = '" & gstrUNITID & "' and h.Doc_No = '" & objDocno & "' "
                        Else
                            strSql = " select 'HDR'+','+'" + STRCUSTVENDCODE + "'+','+'BOSCH EDI-TEAM'+','+'1760'+','+ CAST(h.Doc_No AS VARCHAR(10)) + ','"
                            strSql += " +CAST(DBO.FN_DATEFORMAT_YEARMONTH(INVOICE_DATE) AS VARCHAR(6)) "
                            strSql += " + CASE WHEN LEN(DAY(INVOICE_DATE)) = 1 THEN '0' + CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) ELSE  CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) END"
                            strSql += " + replace(convert(varchar(8),h.ent_dt,108),':', '') +', '+CAST(h.Doc_No AS VARCHAR(10))+','"
                            strSql += " +'1760'+','+'CLJ01'+'" + STRCUSTVENDCODE + ",1760'+','+'M'+','+CAST(h.Doc_No AS VARCHAR(10))+','+d.cust_item_code+','+"
                            strSql += " d.item_code+','+'03'+','+ CASE WHEN LEN(DAY(INVOICE_DATE)) = 1 THEN '0' + CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) ELSE  CAST(DAY(INVOICE_DATE) AS VARCHAR(2)) END + "
                            strSql += " CASE WHEN LEN(Month(INVOICE_DATE)) = 1 THEN '0' + CAST(Month(INVOICE_DATE) AS VARCHAR(2)) ELSE  CAST(Month(INVOICE_DATE) AS VARCHAR(2)) END +"
                            strSql += " CAST(year(INVOICE_DATE) AS VARCHAR(4))+','+ cast(d.sales_quantity as varchar(20))+',PCE,'+CAST(d.cust_ref AS VARCHAR(25))+','+isnull(h.final_destination,'')+"
                            strSql += " '" + STRCUSTVENDCODE + ",5101'  from saleschallan_dtl h inner join sales_dtl d on "
                            strSql += " h.unit_code = d.unit_code And h.doc_no = d.doc_no where h.UNIT_CODE = '" & gstrUNITID & "' and h.Doc_No = '" & objDocno & "' "

                        End If
                        oRDR_Hdr = SqlConnectionclass.ExecuteReader(strSql)
                        strSql = "select 'DTL'+','+cast(rank() over (order by d.item_code) as varchar(2))+', 12, NIL,'+d.item_code+','+cast(d.binquantity as varchar(8))+',PCE'+',S,11111,'+" & _
                            " CAST(h.Doc_No AS VARCHAR(10))   from saleschallan_dtl h inner join sales_dtl d on" & _
                            " h.unit_code = d.unit_code And h.doc_no = d.doc_no" & _
                            " where h.UNIT_CODE = '" & gStrUnitId & "' and h.Doc_No = '" & objDocno & "'"
                        oRDR_Dtl = SqlConnectionclass.ExecuteReader(strSql)
                        If oRDR_Hdr.HasRows Then
                            strSql = "SELECT BOSCHASNLOC FROM SALES_PARAMETER WHERE UNIT_CODE = '" & gStrUnitId & "'"
                            strASNFilepath = SqlConnectionclass.ExecuteScalar(strSql)
                            gstrASNPath = strASNFilepath
                            If Directory.Exists(gstrASNPath) = False Then
                                Directory.CreateDirectory(gstrASNPath)
                            End If
                            strASNFilepath = gstrASNPath & "\" & objDocno.ToString() & ".text"
                            If File.Exists(strASNFilepath) Then
                                File.Delete(strASNFilepath)
                            End If
                            fs = File.Create(strASNFilepath)
                            sw = New StreamWriter(fs)
                            While oRDR_Hdr.Read
                                sw.WriteLine(oRDR_Hdr(0).ToString())
                                oRDR_Dtl.Read()
                                sw.WriteLine(oRDR_Dtl(0).ToString())
                            End While
                        End If
                        gstrASNPath = gstrUserMyDocPath
                        sw.Close()
                        fs.Close()
                        BOSCHASNFileGeneration_MTL = True
                    Else
                        MessageBox.Show("Unable To Get Plant Code For The Customer: " & txtCustomer.Text & " While Generating ASN File." & vbCrLf & _
                                            "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        BOSCHASNFileGeneration_MTL = False
                        Exit Function
                    End If
                End If
            Next
            If intCount = 0 Then
                MessageBox.Show("Select invoice.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                BOSCHASNFileGeneration_MTL = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            FileClose(1)
        End Try
    End Function

    Private Function SMRCASNFileGeneration_MTL() As String
        Dim intCount As Short
        Dim STRPLANTCODE As String
        Dim strSql As String
        Dim strRecord As String
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim dblcummulativeQty As Double
        Dim FSO As New Scripting.FileSystemObject
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim objDocno As Object = Nothing
        Dim countInv As Integer = 0
        Dim objSel As Object = Nothing
        Dim oRDR As SqlDataReader
        Dim sqlCmd As New SqlCommand

        Try
            SMRCASNFileGeneration_MTL = False
            For countInv = 1 To spdInv.MaxRows
                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL : objSel = Nothing : objSel = spdInv.Text
                If objSel = "" Then objSel = 0
                If CDbl(objSel) = 1 Then
                    intCount += 1
                    spdInv.Row = countInv : spdInv.Col = ENUM_ASN.INVOICENO
                    objDocno = Nothing : objDocno = spdInv.Text


                    STRPLANTCODE = Trim(Find_Value("SELECT ISNULL(PLANT_CODE ,'') FROM customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' and customer_code='" & txtCustomer.Text & "'"))
                   
                    If Len(STRPLANTCODE) > 0 Then
                        strSql = ""
                        With sqlCmd
                            .CommandText = "USP_SMRCASNFileGen"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@STRPLANTCODE", SqlDbType.VarChar, 10).Value = STRPLANTCODE
                            .Parameters.Add("@TempDocNo", SqlDbType.Int).Value = objDocno
                        End With
                        oRDR = SqlConnectionclass.ExecuteReader(sqlCmd)
                        strRecord = ""
                        If oRDR.HasRows Then
                            strSql = "SELECT SMRCASNLoc FROM SALES_PARAMETER WHERE UNIT_CODE = '" & gstrUNITID & "'"
                            strASNFilepath = SqlConnectionclass.ExecuteScalar(strSql)
                            gstrASNPath = strASNFilepath
                            If gstrASNPath.Length > 0 Then
                                If Directory.Exists(gstrASNPath) = False Then
                                    Directory.CreateDirectory(gstrASNPath)
                                End If
                            End If
                            strASNFilepath = gstrASNPath & "\" & objDocno.ToString() & ".text"
                            fs = File.Create(strASNFilepath)
                            sw = New StreamWriter(fs)

                            While oRDR.Read
                                sw.WriteLine(oRDR(0).ToString)
                            End While
                        End If
                        gstrASNPath = gstrUserMyDocPath
                        sw.Close()
                        fs.Close()
                        SMRCASNFileGeneration_MTL = True
                    Else
                        MessageBox.Show("Unable To Get Plant Code For The Customer: " & txtCustomer.Text & " While Generating ASN File." & vbCrLf & _
                                            "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        SMRCASNFileGeneration_MTL = False
                        Exit Function
                    End If
                End If
            Next
            If intCount = 0 Then
                MessageBox.Show("Select invoice.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                SMRCASNFileGeneration_MTL = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            FileClose(1)
        End Try
    End Function

    Private Function PKCASNFileGeneration_MTL() As String
        Dim intCount As Short
        Dim STRPLANTCODE As String
        Dim strSql As String
        Dim strRecord As String
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim FSO As New Scripting.FileSystemObject
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim objDocno As Object = Nothing
        Dim countInv As Integer = 0
        Dim objSel As Object = Nothing
        Dim oRDR As SqlDataReader
        Try
            PKCASNFileGeneration_MTL = False
            For countInv = 1 To spdInv.MaxRows
                spdInv.Row = countInv : spdInv.Col = ENUM_ASN.SEL : objSel = Nothing : objSel = spdInv.Text
                If objSel = "" Then objSel = 0
                If CDbl(objSel) = 1 Then
                    intCount += 1
                    spdInv.Row = countInv : spdInv.Col = ENUM_ASN.INVOICENO
                    objDocno = Nothing : objDocno = spdInv.Text
                    STRPLANTCODE = Trim(Find_Value("SELECT ISNULL(PLANT_CODE ,'') FROM customer_mst WHERE UNIT_CODE='" + gStrUnitId + "' and customer_code='" & txtCustomer.Text & "'"))
                    If Len(STRPLANTCODE) > 0 Then
                        strSql = ""
                        strSql = "  select CAST(H.Doc_No AS VARCHAR(10)) + ','+CAST(DBO.FN_DATEFORMAT_YEARMONTH(H.INVOICE_DATE) AS VARCHAR(6)) " & _
                            "+ CASE WHEN LEN(DAY(H.INVOICE_DATE)) = 1 THEN '0' + CAST(DAY(H.INVOICE_DATE) AS VARCHAR(2)) ELSE  CAST(DAY(H.INVOICE_DATE) AS VARCHAR(2)) END" & _
                            "+ replace(convert(varchar(8),H.ent_dt,108),':','') +','+'Picked By Customer Tpt'+','+CAST(I.WEIGHT AS VARCHAR(10))+','+CAST(I.WEIGHT AS VARCHAR(10))" & _
                            "+ '0932' + ',' + H.PORT_OF_DISCHARGE + ',' + '5161' + ',' + 'PKC-DE' + ',' + RTRIM(H.MODE_OF_SHIPMENT) + ',' + RTRIM(H.TRANSPORT_TYPE) + ',' + ''+','+''+','" & _
                            "+ CAST((D.TO_BOX - D.FROM_BOX + 1) AS VARCHAR(4)) + ISNULL(D.PACKING_TYPE,'') + ','+ cast(cast((D.SALES_QUANTITY / (D.TO_BOX - D.FROM_BOX + 1)) as numeric(12,4)) as varchar(12))" & _
                            "+ ',' + D.ITEM_CODE + ',' + D.CUST_ITEM_CODE + ',' + CAST(D.SALES_QUANTITY AS VARCHAR(12)) + ',' + '1'" & _
                            " from saleschallan_dtl h inner join sales_dtl d on h.unit_code = d.unit_code and h.doc_no = d.doc_no " & _
                            " inner join item_mst i on i.unit_code = d.unit_code and i.item_code = d.item_code" & _
                            " where H.UNIT_CODE = '" + gStrUnitId + "' and H.Doc_No = '" & objDocno & "'"
                        oRDR = SqlConnectionclass.ExecuteReader(strSql)
                        strRecord = ""
                        If oRDR.HasRows Then
                            strSql = "SELECT PKCASNLOC FROM SALES_PARAMETER WHERE UNIT_CODE = '" & gStrUnitId & "'"
                            strASNFilepath = SqlConnectionclass.ExecuteScalar(strSql)
                            gstrASNPath = strASNFilepath
                            If gstrASNPath.Length > 0 Then
                                If Directory.Exists(gstrASNPath) = False Then
                                    Directory.CreateDirectory(gstrASNPath)
                                End If
                            End If
                            strASNFilepath = gstrASNPath & "\" & objDocno.ToString() & ".text"
                            fs = File.Create(strASNFilepath)
                            sw = New StreamWriter(fs)

                            While oRDR.Read
                                sw.WriteLine(oRDR(0).ToString)
                            End While
                        End If
                        gstrASNPath = gstrUserMyDocPath
                        sw.Close()
                        fs.Close()
                        PKCASNFileGeneration_MTL = True
                    Else
                        MessageBox.Show("Unable To Get Plant Code For The Customer: " & txtCustomer.Text & " While Generating ASN File." & vbCrLf & _
                                            "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        PKCASNFileGeneration_MTL = False
                        Exit Function
                    End If
                End If
            Next
            If intCount = 0 Then
                MessageBox.Show("Select invoice.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                PKCASNFileGeneration_MTL = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            FileClose(1)
        End Try
    End Function

    Private Sub btnPrintLabels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintLabels.Click
        Try
            If spdInv.MaxRows > 0 Then
                If (gstrUNITID = "MGS" Or gstrUNITID = "VF1") Then
                    If Not ValidateBeforeSave() Then
                        gblnCancelUnload = True
                        gblnFormAddEdit = True
                        Exit Sub
                    Else
                        Call GenerateLabels()
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub GenerateLabels()
        Try
            Dim invoice As String = String.Empty
            Dim count As Int16 = 0
            Dim intLoopCounter As Int16 = 0
            Dim spdVal As Object = Nothing
            Dim SqlAdp = New SqlDataAdapter
            Dim DSLBLDTL = New DataSet
            Dim strBarcodeMsg As String = String.Empty

            With Me.spdInv
                For intLoopCounter = 1 To .MaxRows

                    spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.SEL
                    spdVal = Nothing : spdVal = spdInv.Text
                    If spdVal = "" Then spdVal = 0

                    If CDbl(spdVal) = 1 Then

                        spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.INVOICENO
                        spdVal = Nothing : spdVal = spdInv.Text

                        invoice = invoice + spdVal + ";"
                        count = count + 1

                    End If

                Next
            End With

            If count = 0 Then
                MsgBox("Kindly select the Invoice to print Labels!")
                Exit Sub
            End If

            If count > 0 Then

                Dim sqlCmd As New SqlCommand()

                With sqlCmd
                    .CommandText = "USP_ASNLABELS_PRINT"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Connection = SqlConnectionclass.GetConnection()
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@MODE", "PRINT")
                    .Parameters.AddWithValue("@CUSTOMERCODE", txtCustomer.Text.Trim.ToString())
                    .Parameters.AddWithValue("@UNITCODE", gstrUNITID)
                    .Parameters.AddWithValue("@IPADDRESS", gstrIpaddressWinSck)
                    .Parameters.AddWithValue("@INVOICENO", invoice)
                    SqlAdp.SelectCommand = sqlCmd
                    SqlAdp.Fill(DSLBLDTL)
                    .Dispose()
                End With

                If DSLBLDTL.Tables.Count > 0 Then
                    If DSLBLDTL.Tables(0).Rows.Count > 0 Then
                        For intLoopCounter = 0 To DSLBLDTL.Tables(0).Rows.Count - 1
                            strBarcodeMsg = GenerateBarCode_LINELEVEL_2dbarcode(gstrUserMyDocPath, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("MASTERLABEL").ToString.Trim)

                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                Exit Sub
                            Else
                                If SaveBarCodeImage_singlelevelso_2DBARCODE(gstrUserMyDocPath, DSLBLDTL.Tables(0).Rows(intLoopCounter).Item("MASTERLABEL").ToString.Trim) = False Then
                                    MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                    Exit Sub
                                End If
                            End If
                        Next
                    End If
                End If

                ' here report integration starts

                Dim objReport As ReportDocument
                Dim frmReportViewer As New eMProCrystalReportViewer
                'Dim strSelectionFormula As String = String.Empty

                objReport = frmReportViewer.GetReportDocument()
                frmReportViewer.ShowPrintButton = True
                frmReportViewer.ShowZoomButton = True

                With objReport
                    .Load(My.Application.Info.DirectoryPath & "\Reports\ASNLabels_RSA.rpt")
                    .SetParameterValue("@UNITCODE", gstrUNITID)
                    .SetParameterValue("@IPADDRESS", gstrIpaddressWinSck)
                    
                    '  .RecordSelectionFormula = "{FORD_ASN_INVLABELS_TEMP.UNIT_CODE} = '" & gstrUNITID & "' AND {FORD_ASN_INVLABELS_TEMP.IPADDRESS}='" & gstrIpaddressWinSck & "'"
                End With

                frmReportViewer.Show()

                ' here report integration ends  

            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ' Barcode Methods
    Function GenerateBarCode_LINELEVEL_2dbarcode(ByVal path As Object, ByVal strBarCode As String) As String

        Try

            Dim strPath As String
            Dim fso As New Scripting.FileSystemObject
            Dim objCRUFLQRC As New CRUFLQRC.UFL()
            Dim ts As Object

            strPath = ""

            strPath = path & "BarcodeImg"

            Dim BarcodeImage As Bitmap = create2DImage(GetEncodedString(strBarCode.Trim(), ""))
            BarcodeImage.Save(strPath & ".JPEG")
            If Dir(strPath & ".txt") <> "" Then Kill(strPath & ".txt")
            ts = fso.OpenTextFile(strPath & ".txt", Scripting.IOMode.ForWriting, True)
            ts.Write(strBarCode)
            ts.Close()
            ts = Nothing

            GenerateBarCode_LINELEVEL_2dbarcode = "Y¤"

            Exit Function
        Catch ex As Exception

            GenerateBarCode_LINELEVEL_2dbarcode = "N¤" & Err.Description
            RaiseException(ex)
        End Try
    End Function

    Public Function SaveBarCodeImage_singlelevelso_2DBARCODE(ByVal pstrPath As String, ByVal strbarcodestring As String) As Boolean
        Try
            Dim stimage As ADODB.Stream
            Dim strQuery As String
            Dim Rs As ADODB.Recordset
            SaveBarCodeImage_singlelevelso_2DBARCODE = True
            stimage = New ADODB.Stream
            stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
            stimage.Open()
            pstrPath = pstrPath & "BarcodeImg.JPEG"
            stimage.LoadFromFile(pstrPath)
            strQuery = "select BARCODEIMAGE from FORD_ASN_INVLABELS_TEMP where IPADDRESS='" & Trim(gstrIpaddressWinSck) & "' and UNIT_CODE = '" & gstrUNITID & "' AND MASTERLABEL='" & strbarcodestring & "'"
            Rs = New ADODB.Recordset
            Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            Rs.Fields("BARCODEIMAGE").Value = stimage.Read
            Rs.Update()
            Rs.Close()
            Rs = Nothing
            Exit Function
        Catch ex As Exception
            SaveBarCodeImage_singlelevelso_2DBARCODE = False
            RaiseException(ex)
        End Try
    End Function

    Private Function GetEncodedString(ByVal DataString As String, ByVal DataType As String) As String
        Dim Encoded_1 As String = EncodedString(DataString)
        System.GC.Collect()
        Dim Encoded_2 As String = EncodedString(DataString)

        If Encoded_1 = Encoded_2 Then
            Return Encoded_1
        Else
            For count As Integer = 0 To 4
                Encoded_1 = String.Empty
                Encoded_2 = String.Empty

                Encoded_1 = EncodedString(DataString)
                Encoded_2 = EncodedString(DataString)

                If Encoded_1 = Encoded_2 Then
                    Return Encoded_1
                End If
            Next
        End If

        If Encoded_1 = Encoded_2 Then
            Return Encoded_1
        Else
            Throw New Exception(DataType & " String encoding mismatch.")
        End If
    End Function

    Private Function create2DImage(ByVal data As String) As Bitmap

        Dim barcode As New Bitmap(1, 1)

        Dim PFC As New PrivateFontCollection
        PFC.AddFontFile("c:\windows\fonts\MW6Matrix.TTF")
        Dim FF As FontFamily = PFC.Families(0)
        Dim fontName As New Font(FF, 30)

        Dim graphics__1 As Graphics = Graphics.FromImage(barcode)
        Dim dataSize As SizeF = graphics__1.MeasureString(data, fontName)

        barcode = New Bitmap(barcode, dataSize.ToSize())
        graphics__1 = Graphics.FromImage(barcode)
        graphics__1.Clear(Color.White)
        graphics__1.TextRenderingHint = TextRenderingHint.SingleBitPerPixel

        graphics__1.DrawString(data, fontName, New SolidBrush(Color.Black), 0, 0)
        graphics__1.Flush()
        fontName.Dispose()
        graphics__1.Dispose()

        Return barcode
    End Function

    Private Function EncodedString(ByVal DataString As String) As String
        Dim objCRUFLQRC As New CRUFLQRC.UFL()
        Return objCRUFLQRC.MW6Encoder(DataString, 0, 0, 0)
    End Function

    ' Barcode Methods

    Private Sub btn_PrintSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_PrintSummary.Click
        Try
            If spdInv.MaxRows > 0 Then
                If (gstrUNITID = "MGS" Or gstrUNITID = "VF1") Then
                    If Not ValidateBeforeSave() Then
                        gblnCancelUnload = True
                        gblnFormAddEdit = True
                        Exit Sub
                    Else
                        Call GenerateSummary()
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub GenerateSummary()
        Try

            Dim invoice As String = String.Empty
            Dim count As Int16 = 0
            Dim intLoopCounter As Int16 = 0
            Dim spdVal As Object = Nothing
            Dim SqlAdp = New SqlDataAdapter
            Dim DSLBLDTL = New DataSet
            Dim strBarcodeMsg As String = String.Empty
            Dim strasnconveyancestring As String = String.Empty
            Dim spdVal_trip As Object = Nothing
            Dim spdVal_triptype As Object = Nothing
            Dim strManualTrip As String = String.Empty

            If (Len(txtTripNo.Text) <> 7) Then
                MsgBox("Please enter Manual Delivery Note Number of length 7.", MsgBoxStyle.Information, ResolveResString(10059))
                Exit Sub
            End If

            With Me.spdInv
                For intLoopCounter = 1 To .MaxRows

                    spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.SEL
                    spdVal = Nothing : spdVal = spdInv.Text
                    If spdVal = "" Then spdVal = 0

                    If CDbl(spdVal) = 1 Then

                        spdInv.Row = intLoopCounter : spdInv.Col = ENUM_ASN.INVOICENO
                        spdVal = Nothing : spdVal = spdInv.Text

                        invoice = invoice + spdVal + ";"
                        count = count + 1

                        Exit For
                    End If

                Next
            End With

            If count = 0 Then
                MsgBox("Kindly select the Invoice to print Summary!")
                Exit Sub
            End If

            strManualTrip = txtTripNo.Text

            If mblnASnConveyance = True Then
                'If spdtrip.MaxRows > 0 Then

                '    For intLoopCounter = 1 To spdtrip.MaxRows
                '        spdtrip.Row = intLoopCounter : spdtrip.Col = ENUM_TRIP.SEL
                '        spdVal_trip = Nothing : spdVal_trip = spdtrip.Text
                '        If spdVal_trip = "" Then spdVal_trip = 0

                '        If CDbl(spdVal_trip) = 1 Then
                '            spdtrip.Row = intLoopCounter : spdtrip.Col = ENUM_TRIP.TRIPNO
                '            spdVal_trip = Nothing : spdVal_trip = spdtrip.Text
                '            Exit For
                '        End If
                '    Next
                'End If

                'If spdtriptype.MaxRows > 0 Then
                '    For intLoopCounter = 1 To spdtriptype.MaxRows
                '        spdtriptype.Row = intLoopCounter : spdtriptype.Col = ENUM_TRIPTYPE.SEL
                '        spdVal_triptype = Nothing : spdVal_triptype = spdtriptype.Text
                '        If spdVal_triptype = "" Then spdVal_triptype = 0

                '        If CDbl(spdVal_triptype) = 1 Then
                '            spdtriptype.Row = intLoopCounter : spdtriptype.Col = ENUM_TRIPTYPE.TRIPTYPE
                '            spdVal_triptype = Nothing : spdVal_triptype = spdtriptype.Text
                '            Exit For
                '        End If
                '    Next
                'End If


                strasnconveyancestring = Find_Value("SELECT DBO.UDF_GETASNCONVEYANCE('" & gstrUNITID & "','" & txtCustomer.Text.Trim.ToString() & "','" & spdVal_trip & "','" & spdVal_triptype & "')")

            End If

            If count > 0 Then
                ' here report integration starts

                Dim objReport As ReportDocument
                Dim frmReportViewer As New eMProCrystalReportViewer
                'Dim strSelectionFormula As String = String.Empty

                objReport = frmReportViewer.GetReportDocument()
                frmReportViewer.ShowPrintButton = True
                frmReportViewer.ShowZoomButton = True

                With objReport
                    .Load(My.Application.Info.DirectoryPath & "\Reports\HdrLblDesign.rpt")
                    .SetParameterValue("@UNITCODE", gstrUNITID)
                    .SetParameterValue("@INVOICENO", invoice)
                    .SetParameterValue("@CUSTOMERCODE", txtCustomer.Text.Trim.ToString())
                    .SetParameterValue("@MODE", "PRINT_SUMM")
                    .SetParameterValue("@CONVEYANCE", strasnconveyancestring)
                    .SetParameterValue("@ManualTrip", strManualTrip)
                End With

                frmReportViewer.Show()

                ' here report integration ends  

            End If


        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtTripNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTripNo.Validating
        Try
            Dim val As Decimal = 0.0
            If Not Decimal.TryParse(txtTripNo.Text.Trim(), val) Then
                MessageBox.Show("Enter valid Trip Desc in numeric only.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtTripNo.Text = ""
                txtTripNo.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub lbltripselection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbltripselection.Click

    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub
End Class
