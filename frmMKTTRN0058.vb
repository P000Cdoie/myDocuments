Option Strict Off
Option Explicit On
Imports System.Text
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0058
    Inherits System.Windows.Forms.Form
    '****************************************************
    'Copyright (c)  -  MIND
    'Name of module -  FRMMKTTRN0058.frm
    'Created By     -  Shubhra Verma
    'Created On     -  20 Oct 2007
    'Issue ID       -  21466
    'description    -  Sales Order And Schedule Uploading
    'Revised By     -  Manoj Kr Vaish
    'Revised date   -  02 Jul 2008 (eMpro-20080702-20004)
    'Revised History-  While Authorizing the SO the dates(Valid date & effective Date) was updating wrong 
    'Revised By     -  Manoj Kr Vaish
    'Revised date   -  05 Dec 2008 (eMpro-20081205-24448)
    'Revised History-  While Authorizing the CSP Amount is saving with 0 value in Sales Order Table
    'Revised By     -  Manoj Kr Vaish
    'Revised date   -  05 Jan 2009 (eMpro-20090105-25543)
    'Revised History-  While Authorizing the Sales order conversion from varchar to numeric is coming
    'Revised By     -  Shubhra Verma
    'Revised On     -  20 Jan 2009 to 21 Jan 2009
    'Issue ID       -  eMpro-20090120-26215
    'Revised History-  If Schedule is for current Month, then system should divide that required qty in remaining
    '                  working days of that month.
    '                  if some Qty is already despatched for the item, then system should use New Schedule Qty - 
    '                  Despatched Qty as New Schedule Qty
    '********************************************************************************************************
    'Revised By     :   Shubhra Verma
    'Revised On     :   06 Jun 2011
    'Description    :   Multi Unit Changes
    '********************************************************************************************************
    'Revised By     -  Prashant Rajpal
    'Revised On     -  01 Aug 2011
    'Issue ID       -  10118036
    'Revised History-  VAT Is not calculated in Sales Order Table 
    '********************************************************************************************************
    'MODIFIED BY VIRENDRA GUPTA 0N 09 NOV 2011 FOR CHANGE MANAGEMENT
    '********************************************************************************************************
    'Modified By Shubhra on 31 Mar 2012 to increase the length of Doc No
    '**********************************************************************************
    'Revised By      : Prashant Rajpal
    'Issue ID        : 10624527
    'Revision Date   : 18-SEP-2014
    'History         : Check TOOL COST AND DISCOUNT VALUE IN SO AND SCHEDULE UPLOADING FORM 
    '****************************************************************************************
    'Revised By      : Abhinav Kumar
    'Issue ID        : 10709130
    'Revision Date   : 20-NOV-2014
    'History         : Issues in SO uploading form
    '****************************************************************************************
    'Revised By      : Prashant Rajpal
    'Issue ID        : 10727228 
    'Revision Date   : 17 dec 2014
    'History         : Issues in SO uploading form
    '****************************************************************************************
    '***************************************************************************************
    'REVISED BY         :   Prashant rajpal 
    'ISSUE ID           :   101528442 — eMPro : SO and Schedule uploading
    'DESCRIPTION        :   eMPro : SO and Schedule uploading
    'REVISION DATE      :   25 May 2018
    '***************************************************************************************

    Dim m_strCustomerCode As String
    Dim mintFormIndex As Short
    Dim StrDocNum As String
    Dim mobjEmpDll As New EMPDataBase.EMPDB(gstrUNITID)
    Dim mrsEmpDll As New EMPDataBase.CRecordset
    Dim varItemCode As Object
    Dim mbln_View As Boolean
    Dim Flag As Boolean
    Dim blnlinelevelcustomer As Boolean = False
    Dim mblnopensalesorder As Boolean = False
    Dim mblnUPDATE_DMS_SO_SCH_UPLD As Boolean = True
    Dim mblnSOUPLD_ForexCurrency As Boolean = False

    Private Enum ENUM_Grid
        salesorder = 1
        InternalPartNo
        CustPartNo
        ItemDesc
        CustPartDesc
        Qty
        DespatchQty
        PrevQty
        ExWorks
        custsupply
        ToolCost
        EDPer
        AddEx
        Cess
        SHECess
        VAT
        discount_type
        discount_value
        HSNSACCODE
        CGST_TAX
        SGST_TAX
        IGST_TAX
        COMP_ECC
        PACKING_PER
        Remarks
    End Enum
    Private Sub FN_Spread_Settings()
        Dim Col As Short
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With SSPurOrd

            .MaxCols = ENUM_Grid.Remarks
            .MaxRows = 0
            .Row = 0
            .Col = ENUM_Grid.salesorder : .Text = "Sales Order"
            .Col = ENUM_Grid.InternalPartNo : .Text = "Int.PartNo"
            .Col = ENUM_Grid.CustPartNo : .Text = "CustPartNo"
            .Col = ENUM_Grid.ItemDesc : .Text = "Item Desc"
            .Col = ENUM_Grid.CustPartDesc : .Text = "CustPart Desc"
            .Col = ENUM_Grid.Qty : .Text = "Qty"
            .Col = ENUM_Grid.DespatchQty : .Text = "Despatch Qty"
            .Col = ENUM_Grid.PrevQty : .Text = "Previous Qty"
            .Col = ENUM_Grid.ExWorks : .Text = "Rate"
            .Col = ENUM_Grid.custsupply : .Text = "Cust Supply"
            .Col = ENUM_Grid.ToolCost : .Text = "Tool Cost"
            .Col = ENUM_Grid.EDPer : .Text = "E.D.%"
            .Col = ENUM_Grid.AddEx : .Text = "AddE.D.%"
            .Col = ENUM_Grid.SHECess : .Text = "SHECess%"
            .Col = ENUM_Grid.Cess : .Text = "Cess%"
            .Col = ENUM_Grid.VAT : .Text = "VAT%"
            .Col = ENUM_Grid.discount_type : .Text = "DISCOUNT TYPE"
            .Col = ENUM_Grid.discount_value : .Text = "DISCOUNT PER"
            .Col = ENUM_Grid.HSNSACCODE : .Text = "HSN/SACCODE"
            .Col = ENUM_Grid.CGST_TAX : .Text = "CGST"
            .Col = ENUM_Grid.SGST_TAX : .Text = "SGST"
            .Col = ENUM_Grid.IGST_TAX : .Text = "IGST"
            .Col = ENUM_Grid.COMP_ECC : .Text = "COMP CC"
            .Col = ENUM_Grid.PACKING_PER : .Text = "PACKING(PER UNIT)"
            .Col = ENUM_Grid.Remarks : .Text = "Remarks"
            'gst changes

            .Row = -1
            .Col = ENUM_Grid.salesorder
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.InternalPartNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.CustPartNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.ItemDesc
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.CustPartDesc
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.Qty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.DespatchQty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.PrevQty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.ExWorks
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .Col = ENUM_Grid.custsupply
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.ToolCost
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.EDPer
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.AddEx
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.SHECess
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.Cess
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.VAT
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.discount_type
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = ENUM_Grid.discount_value
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Col = ENUM_Grid.HSNSACCODE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.CGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.SGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.IGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.PACKING_PER
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.Remarks
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .set_RowHeight(0, 20)
            .set_ColWidth(ENUM_Grid.salesorder, 10)
            .set_ColWidth(ENUM_Grid.InternalPartNo, 10)
            .set_ColWidth(ENUM_Grid.CustPartNo, 10)
            .set_ColWidth(ENUM_Grid.ItemDesc, 0)
            .set_ColWidth(ENUM_Grid.CustPartDesc, 0)
            .set_ColWidth(ENUM_Grid.Qty, 5)
            .set_ColWidth(ENUM_Grid.DespatchQty, 8)
            .set_ColWidth(ENUM_Grid.PrevQty, 8)
            .set_ColWidth(ENUM_Grid.custsupply, 5)
            .set_ColWidth(ENUM_Grid.ToolCost, 5)
            .set_ColWidth(ENUM_Grid.ExWorks, 5)
            .set_ColWidth(ENUM_Grid.EDPer, 5)
            .set_ColWidth(ENUM_Grid.AddEx, 5)
            .set_ColWidth(ENUM_Grid.SHECess, 5)
            .set_ColWidth(ENUM_Grid.Cess, 5)
            .set_ColWidth(ENUM_Grid.VAT, 5)
            .set_ColWidth(ENUM_Grid.discount_type, 7)
            .set_ColWidth(ENUM_Grid.discount_value, 7)
            .set_ColWidth(ENUM_Grid.HSNSACCODE, 14)
            .set_ColWidth(ENUM_Grid.CGST_TAX, 7)
            .set_ColWidth(ENUM_Grid.SGST_TAX, 7)
            .set_ColWidth(ENUM_Grid.COMP_ECC, 7)
            .set_ColWidth(ENUM_Grid.PACKING_PER, 6)
            .Col = ENUM_Grid.Remarks

            .set_ColWidth(ENUM_Grid.Remarks, 20)
            .Row = 1
            .Row = .MaxRows
            .Col = ENUM_Grid.salesorder
            .Col2 = .MaxCols
            .Lock = True
            .BlockMode = True
        End With
        With SpdSOSch
            .MaxCols = 10
            .MaxRows = 4
            .Row = -1
            For Col = 1 To .MaxCols
                .Col = Col
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .set_ColWidth(Col, 5)
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Lock = False
            Next
            .Col = 1
            .Row = 1
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Row = 4
            For Col = 3 To .MaxCols
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Lock = True
            Next
            .RowHeaderDisplay = FPSpreadADO.HeaderDisplayConstants.DispNumbers
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Description, Err.Source, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmbSearch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbSearch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub cmdCancel_Click()
        Dim YesNo As String
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        YesNo = CStr(MsgBox("Do you want to Refresh the changes?", MsgBoxStyle.YesNo, ResolveResString(100)))
        If YesNo = CStr(MsgBoxResult.Yes) Then
            Call RefreshForm()
        End If
        If YesNo = CStr(MsgBoxResult.No) Then
            Exit Sub
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DisplayDailyMktSchedule(ByRef Click_Renamed As Boolean, ByRef NewRow As Integer)
        Dim varQty As Object
        Dim SCHQTY As Integer
        Dim Col, Row As Short
        Dim rsWkngDays As New ClsResultSetDB
        Dim fraction As Short
        Dim rsMinQty As New ClsResultSetDB
        Dim rsSch As New ClsResultSetDB
        Dim varItemCode As Object
        Dim varDespatchQty As Object = Nothing
        Dim varPrevQty As Object = Nothing
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        txtQty.Text = CStr(0)
        txtSchQty.Text = CStr(0)
        rsWkngDays.GetResult("select distinct work_flg,dt from" &
            " calendar_mfg_mst where UNIT_CODE = '" & gstrUNITID & "' AND" &
            " month (dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "'" & " and YEAR (dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' and dt > =  CONVERT(DATETIME,CONVERT(VARCHAR(10),getdate(),103),103) order by dt")
        If rsWkngDays.GetNoRows > 0 Then
            rsWkngDays.MoveFirst()
            For Row = 1 To 4.0#
                For Col = 1 To 10
                    If Row = 1 And Col = 1 Then
                        SpdSOSch.Col = Col
                        SpdSOSch.Row = Row
                        SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        GoTo skip
                    End If
                    If rsWkngDays.EOFRecord = False Then
                        If CStr(Row - 1) + CStr(Col - 1) < DatePart(DateInterval.Day, rsWkngDays.GetValue("DT")) Then
                            SpdSOSch.Col = Col
                            SpdSOSch.Row = Row
                            SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            GoTo skip
                        End If
                    End If
                    SpdSOSch.Col = Col
                    SpdSOSch.Row = Row
                    If rsWkngDays.EOFRecord = False Then
                        If rsWkngDays.GetValue("work_flg") = True Then
                            SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Else
                            SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                        End If
                        rsWkngDays.MoveNext()
                    Else
                        SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    End If
skip:
                Next
            Next
        Else
            For Row = 1 To 4
                For Col = 1 To 10
                    SpdSOSch.Col = Col
                    SpdSOSch.Row = Row
                    SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                Next
            Next
        End If
        SpdSOSch.Col = 1
        SpdSOSch.Col2 = 1
        SpdSOSch.Row = 4
        SpdSOSch.Row2 = 4
        If (SpdSOSch.BackColor) = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) Then
            SpdSOSch.BlockMode = True
            SpdSOSch.Lock = False
            SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            SpdSOSch.BlockMode = False
        End If
        For Row = 1 To 4
            For Col = 1 To 10
                SpdSOSch.SetText(Col, Row, "")
            Next
        Next
        rsWkngDays.ResultSetClose()
        With SSPurOrd
            If Click_Renamed = True Then
                varQty = Nothing
                .GetText(ENUM_Grid.Qty, .ActiveRow, varQty)
                varDespatchQty = Nothing
                SSPurOrd.GetText(ENUM_Grid.DespatchQty, SSPurOrd.ActiveRow, varDespatchQty)
                varPrevQty = Nothing
                SSPurOrd.GetText(ENUM_Grid.PrevQty, SSPurOrd.ActiveRow, varPrevQty)
            Else
                varQty = Nothing
                .GetText(ENUM_Grid.Qty, NewRow, varQty)
                varDespatchQty = Nothing
                SSPurOrd.GetText(ENUM_Grid.DespatchQty, NewRow, varDespatchQty)
                varPrevQty = Nothing
                SSPurOrd.GetText(ENUM_Grid.PrevQty, NewRow, varPrevQty)
            End If
            If Val(Replace(varQty, ",", "")) + Val(varPrevQty) - Val(varDespatchQty) > 0 Then  'eMpro-20090120-26215
                txtQty.Text = Val(Replace(varQty, ",", "")) + Val(varPrevQty) - Val(varDespatchQty)
            Else
                txtQty.Text = CStr(0)
            End If
        End With
        With SpdSOSch
            If Click_Renamed = True Then
                varItemCode = Nothing
                SSPurOrd.GetText(ENUM_Grid.InternalPartNo, SSPurOrd.ActiveRow, varItemCode)
            Else
                varItemCode = Nothing
                SSPurOrd.GetText(ENUM_Grid.InternalPartNo, NewRow, varItemCode)
            End If
            mP_Connection.Execute("set dateformat 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            rsSch.GetResult("select schedule_quantity,trans_date from dailymktschedule_temp" &
                " where UNIT_CODE = '" & gstrUNITID & "' AND item_code = '" & varItemCode & "'" &
                " and month(trans_date) = month('" & getDateForDB(dtpSOdate.Value) & "')" &
                " and year(trans_date) = year('" & getDateForDB(dtpSOdate.Value) & "') AND DOC_NO = '" & txtDocNo.Text & "' AND FILETYPE = 'SOUPLD' order by trans_date")
            If rsSch.GetNoRows > 0 Then
                rsSch.MoveFirst()
                For Row = 1 To 4
                    For Col = 1 To 10
                        If rsSch.EOFRecord Then
                            Exit For
                        End If
                        If Convert.ToDateTime(rsSch.GetValue("trans_date")).Day.ToString.Length = 1 Then
                            If Convert.ToDateTime(rsSch.GetValue("trans_date")).Day = Col - 1 Then
                                SpdSOSch.SetText(Col, Row, rsSch.GetValue("schedule_quantity"))
                                rsSch.MoveNext()
                            End If
                        Else
                            If CDbl(VB.Left(CStr(Convert.ToDateTime(rsSch.GetValue("trans_date")).Day), 1)) = Row - 1 And CDbl(VB.Right(CStr(Convert.ToDateTime(rsSch.GetValue("trans_date")).Day), 1)) = Col - 1 Then
                                SpdSOSch.SetText(Col, Row, rsSch.GetValue("schedule_quantity"))
                                rsSch.MoveNext()
                            End If
                        End If
                    Next
                Next
            End If
            rsSch.ResultSetClose()
        End With
        For Row = 1 To SpdSOSch.MaxRows
            For Col = 1 To SpdSOSch.MaxCols
                SpdSOSch.Row = Row
                SpdSOSch.Col = Col
                If SpdSOSch.Value = "" Then
                    SCHQTY = SCHQTY + 0
                Else
                    SCHQTY = SCHQTY + SpdSOSch.Value
                End If
            Next
        Next
        txtSchQty.Text = CStr(SCHQTY)
        txtDiff.Text = CStr(Val(txtQty.Text) - Val(txtSchQty.Text))
        txtSchQty.Text = CStr(SCHQTY)
        txtDiff.Text = CStr(Val(txtQty.Text) - Val(txtSchQty.Text))
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdClose_Click()
        On Error GoTo ErrHandler
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustHelp.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select customer_Code,cust_name from customer_mst where unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", "Help", 2)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtCustomerCode.Text = strHelp(0)
                    Me.LblCustomerName.Text = strHelp(1)
                    Me.txtCustomerCode.Focus()
                End If
            Else
                Me.txtCustomerCode.Text = ""
                Me.LblCustomerName.Text = ""
                Me.txtCustomerCode.Focus()
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdExDutyHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExDutyHelp.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate Where unit_code='" & gstrUNITID & "'", "Excise Duty", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtExciseDuty.Text = strHelp(0)
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdFileHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFileHelp.Click
        On Error GoTo ErrHandler
        CommanDLogOpen.InitialDirectory = gstrLocalCDrive
        CommanDLogOpen.Filter = "Microsoft Excel File (*.xls)|*.xls"
        CommanDLogOpen.ShowDialog()
        Me.txtFileName.Text = CommanDLogOpen.FileName
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelpItem_Click()
        On Error GoTo ErrHandler
        Dim strHelp() As String
        Dim Row, Col As Short
        Dim varCustPart As Object
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, " Select item_code ,item_desc ,cust_drgno," &
           " drg_desc from custitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomerCode.Text & "'" &
           " AND ACTIVE = 1")
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) = "" Then
                    Exit Sub
                End If
                lblItemDesc.Text = strHelp(1)
                lblCustPartDesc.Text = strHelp(3)
                With SSPurOrd
                    For Row = 1 To .MaxRows
                        Col = ENUM_Grid.CustPartNo
                        varCustPart = Nothing
                        .GetText(Col, Row, varCustPart)
                        If strHelp(2) = varCustPart Then
                            .Row = Row
                            .Col = ENUM_Grid.CustPartNo
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Exit Sub
                        End If
                    Next
                End With
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function ValidatebeforeSave() As Boolean
        Dim rsSoUpld As New ClsResultSetDB
        Dim strQry As String
        Dim SALESORDERLIST As String
        Dim STROPENCLOSEDSO As String


        If mblnopensalesorder = 0 Then
            STROPENCLOSEDSO = 0
        Else
            STROPENCLOSEDSO = 1
        End If

        rsSoUpld.GetResult("SELECT * FROM so_upld_dtl Where Doc_No = '" & txtDocNo.Text & "' and UNIT_CODE = '" & gstrUNITID & "'")
        If rsSoUpld.GetNoRows > 0 Then
            rsSoUpld.MoveFirst()
            While Not rsSoUpld.EOFRecord
                'SALESORDERLIST = Find_Value("Select dbo.UDF_CHECK_NOOF_ACTIVESO('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & rsSoUpld.GetValue("salesorder") & "','" & rsSoUpld.GetValue("partno") & "','" & rsSoUpld.GetValue("item_code") & "',CONVERT(VARCHAR(10),Convert(varchar(10),'" & dtpDateFrom.Value.ToString & "'),113))")
                SALESORDERLIST = Find_Value("Select dbo.UDF_CHECK_NOOF_OPEN_CLOSED_SALESORDER('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtCustomerCode.Text & "','" & rsSoUpld.GetValue("salesorder") & "','" & rsSoUpld.GetValue("partno") & "','" & rsSoUpld.GetValue("item_code") & "','" & STROPENCLOSEDSO & "')")
                If SALESORDERLIST <> "" Then
                    If Len(SALESORDERLIST) >= 1 Then
                        MsgBox("Already one sales order is Active for item code ." & rsSoUpld.GetValue("item_code") & vbCrLf & " Sales Order Details : " & SALESORDERLIST, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                        ValidatebeforeSave = True
                        Exit Function
                    End If
                End If
                rsSoUpld.MoveNext()
            End While
        End If
        rsSoUpld.ResultSetClose()

    End Function
    Private Sub cmdAuthorize_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdAuthorize.ButtonClick
        On Error GoTo Errorhandler
        'Dim rsInsert As New ClsResultSetDB
        Dim rsDlyMkt As New ClsResultSetDB
        Dim rsDup As New ClsResultSetDB
        Dim STRAMENDMENTNO As String
        Dim strsql As String
        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE
                If Len(LTrim(RTrim(txtDocNo.Text))) > 0 Then
                    If ValidatebeforeSave() = True Then Exit Sub
                    Dim mblnCT2Reqd As Boolean
                    Dim mstrshipaddcode As String
                    Dim mstrshipadddesc As String

                    strsql = "SELECT CT2_Reqd_InSO FROM SO_Upld_hdr WHERE UNIT_CODE='" & gstrUNITID & "' and doc_no='" & txtDocNo.Text & "'"
                    mblnCT2Reqd = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql))

                    strsql = "SELECT SHIPADDRESS_CODE FROM SO_Upld_hdr WHERE UNIT_CODE='" & gstrUNITID & "' and doc_no='" & txtDocNo.Text & "'"
                    mstrshipaddcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(strsql))

                    strsql = "SELECT SHIPADDRESS_DESC FROM SO_Upld_hdr WHERE UNIT_CODE='" & gstrUNITID & "' and doc_no='" & txtDocNo.Text & "'"
                    mstrshipadddesc = Convert.ToString(SqlConnectionclass.ExecuteScalar(strsql))

                    mblnSOUPLD_ForexCurrency = Find_Value("SELECT SOUPLD_FOREXCURRENCY FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")

                    mP_Connection.BeginTrans()

                    If blnlinelevelcustomer = False Then
                        STRAMENDMENTNO = GetServerDateTime().ToString("ddMMyy-HHmmss")
                        If UCase(Trim(gstrUNITID)) = "SMC" Or UCase(Trim(gstrUNITID)) = "SMP" Or UCase(Trim(gstrUNITID)) = "MS1" Then
                            strsql = "SELECT TOP 1 1 FROM CUST_ORD_HDR CH,CUST_ORD_DTL CD ,SO_UPLD_DTL SD WHERE CH.UNIT_CODE = CD.UNIT_CODE AND CH.CUST_REF =CD.CUST_REF "
                            strsql += " AND CH.AMENDMENT_NO=CD.AMENDMENT_NO AND SD.UNIT_CODE=CD.UNIT_CODE AND SD.SALESORDER=CD.CUST_REF "
                            'AND SD.ITEM_CODE=CD.ITEM_CODE 
                        '    strsql += " AND SD.PARTNO=CD.CUST_DRGNO "
                            strsql += " AND SD.CUST_CODE=CD.ACCOUNT_CODE AND SD.DOC_NO='" & txtDocNo.Text & "'"
                            strsql += " AND CH.UNIT_CODE = '" & gstrUNITID & "' AND CH.ACCOUNT_CODE='" & txtCustomerCode.Text & "'"
                            strsql += " AND CH.AUTHORIZED_FLAG=1 "
                        Else
                            strsql = "SELECT TOP 1 1 FROM CUST_ORD_HDR CH,CUST_ORD_DTL CD ,SO_UPLD_DTL SD WHERE CH.UNIT_CODE = CD.UNIT_CODE AND CH.CUST_REF =CD.CUST_REF "
                            strsql += " AND CH.AMENDMENT_NO=CD.AMENDMENT_NO AND SD.UNIT_CODE=CD.UNIT_CODE AND SD.SALESORDER=CD.CUST_REF AND SD.ITEM_CODE=CD.ITEM_CODE AND SD.PARTNO=CD.CUST_DRGNO"
                            strsql += " AND SD.CUST_CODE=CD.ACCOUNT_CODE AND SD.DOC_NO='" & txtDocNo.Text & "'"
                            strsql += " AND CH.UNIT_CODE = '" & gstrUNITID & "' AND CH.ACCOUNT_CODE='" & txtCustomerCode.Text & "'"
                            strsql += " AND CH.AUTHORIZED_FLAG=1 AND CH.PO_TYPE NOT IN('E')  "
                        End If


                        If DataExist(strsql) = True Then
                            If UCase(Trim(gstrUNITID)) = "SMC" Or UCase(Trim(gstrUNITID)) = "SMP" Then
                                mP_Connection.Execute(" UPDATE CH SET CH.VALID_DATE=dateadd(d,-1,CONVERT(VARCHAR(10),Convert(varchar(10),'" & dtpDateFrom.Value.ToString & "'),113)) FROM CUST_ORD_HDR CH,CUST_ORD_DTL CD ,SO_UPLD_DTL SD WHERE CH.UNIT_CODE = CD.UNIT_CODE AND CH.CUST_REF =CD.CUST_REF " &
                          " AND CH.AMENDMENT_NO=CD.AMENDMENT_NO AND SD.UNIT_CODE=CD.UNIT_CODE AND SD.SALESORDER=CD.CUST_REF AND SD.PARTNO=CD.CUST_DRGNO " &
                          " AND SD.CUST_CODE=CD.ACCOUNT_CODE AND SD.DOC_NO='" & txtDocNo.Text & "' AND CH.ACTIVE_FLAG ='A'  AND CD.ACTIVE_FLAG ='A' AND CH.VALID_DATE >=CONVERT(VARCHAR(10),Convert(varchar(10),'" & dtpDateFrom.Value.ToString & "'),113) " &
                          " AND EFFECT_DATE <= CONVERT(VARCHAR(10),Convert(varchar(10),'" & dtpDateFrom.Value.ToString & "'),113) AND CH.UNIT_CODE = '" & gstrUNITID & "' AND CH.ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                          " AND CH.AUTHORIZED_FLAG=1 AND CH.PO_TYPE NOT IN('E') ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute(" UPDATE CH SET CH.VALID_DATE=dateadd(d,-1,CONVERT(VARCHAR(10),Convert(varchar(10),'" & dtpDateFrom.Value.ToString & "'),113)) FROM CUST_ORD_HDR CH,CUST_ORD_DTL CD ,SO_UPLD_DTL SD WHERE CH.UNIT_CODE = CD.UNIT_CODE AND CH.CUST_REF =CD.CUST_REF " &
                          " AND CH.AMENDMENT_NO=CD.AMENDMENT_NO AND SD.UNIT_CODE=CD.UNIT_CODE AND SD.SALESORDER=CD.CUST_REF AND SD.ITEM_CODE=CD.ITEM_CODE AND SD.PARTNO=CD.CUST_DRGNO" &
                          " AND SD.CUST_CODE=CD.ACCOUNT_CODE AND SD.DOC_NO='" & txtDocNo.Text & "'AND CH.ACTIVE_FLAG ='A'  AND CD.ACTIVE_FLAG ='A' AND CH.VALID_DATE >=CONVERT(VARCHAR(10),Convert(varchar(10),'" & dtpDateFrom.Value.ToString & "'),113) " &
                          " AND EFFECT_DATE <= CONVERT(VARCHAR(10),Convert(varchar(10),'" & dtpDateFrom.Value.ToString & "'),113) AND CH.UNIT_CODE = '" & gstrUNITID & "' AND CH.ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                          " AND CH.AUTHORIZED_FLAG=1 AND CH.PO_TYPE NOT IN('E') ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If

                            If mblnopensalesorder = True Then
                                If mblnSOUPLD_ForexCurrency = True Or gstrUNITID = "MS1" Then
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                    " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                    " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                    " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                    " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                    " SELECT distinct CUST_CODE,SalesOrder,'" & STRAMENDMENTNO & "',SO_DATE,'','A',CURRENCY_CODE," &
                                    " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                    " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                    " '" & mP_User & "','1'," & " REMARKS,'E',VAT,'" & GetServerDate() & "',ENT_UID,'" & GetServerDate() & "','" & mP_User & "'," &
                                    " '1','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                    " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                Else
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                    " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                    " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                    " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                    " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                    " SELECT distinct CUST_CODE,SalesOrder,'" & STRAMENDMENTNO & "',SO_DATE,'','A',CURRENCY_CODE," &
                                    " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                    " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                    " '" & mP_User & "','1'," & " REMARKS,'O',VAT,'" & GetServerDate() & "',ENT_UID,'" & GetServerDate() & "','" & mP_User & "'," &
                                    " '1','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                    " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                End If
                                '10624527
                                mP_Connection.Execute("INSERT INTO cust_ord_dtl (UNIT_CODE,Account_Code,Packing_Type,Cust_ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag," &
                                  " Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId," &
                                    " OpenSO,SalesTax_Type,PerValue,RevisionNo,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,DISCOUNT_TYPE,DISCOUNT_VALUE,HSNSACCODE," &
                                    " CGSTTXRT_TYPE,SGSTTXRT_TYPE,IGSTTXRT_TYPE ,COMPENSATION_CESS )" &
                                    " SELECT '" & gstrUNITID & "',CUST_CODE,'PKT0',SalesOrder,'" & STRAMENDMENTNO & "',ITEM_CODE,EXWORKS,QUANTITY,0,'A'," &
                                    "cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,packing_per,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid," &
                                    " getdate(),'" & mP_User & "','1',VAT,'1','0','1','1',AddEx ,DISCOUNT_TYPE,DISCOUNT_VALUE ,HSNSACCODE,CGST_TAX,SGST_TAX,IGST_TAX,COMPCC_TAX FROM SO_Upld_dtl" &
                                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                '10624527
                            Else
                                If mblnSOUPLD_ForexCurrency = True Then
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                     " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                     " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                     " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                     " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                     " SELECT distinct CUST_CODE,SalesOrder,'" & STRAMENDMENTNO & "',SO_DATE,'','A',CURRENCY_CODE," &
                                     " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                     " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                     " '" & mP_User & "','1'," & " REMARKS,'E',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                                     " '0','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                     " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                      " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                Else
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                     " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                     " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                     " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                     " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                     " SELECT distinct CUST_CODE,SalesOrder,'" & STRAMENDMENTNO & "',SO_DATE,'','A',CURRENCY_CODE," &
                                     " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                     " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                     " '" & mP_User & "','1'," & " REMARKS,'O',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                                     " '0','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                     " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                      " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                End If

                                mP_Connection.Execute("INSERT INTO cust_ord_dtl (UNIT_CODE,Account_Code,Packing_Type,Cust_ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag," &
                                     " Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId," &
                                     " OpenSO,SalesTax_Type,PerValue,RevisionNo,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,DISCOUNT_TYPE,DISCOUNT_VALUE,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,IGSTTXRT_TYPE ,COMPENSATION_CESS )" &
                                     " SELECT '" & gstrUNITID & "',CUST_CODE,'PKT0',SalesOrder,'" & STRAMENDMENTNO & "',item_code,exworks,Quantity,0,'A'," &
                                     " cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,packing_per,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid," &
                                     " getdate(),'" & mP_User & "','0',VAT,'1','0','1','1',AddEx ,DISCOUNT_TYPE,DISCOUNT_VALUE,HSNSACCODE,CGST_TAX,SGST_TAX,IGST_TAX,COMPCC_TAX   FROM SO_Upld_dtl" &
                                     " WHERE UNIT_CODE = '" & gstrUNITID & "' and DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If

                        Else
                            If mblnopensalesorder = True Then
                                If mblnSOUPLD_ForexCurrency = True Then
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                " SELECT distinct CUST_CODE,SalesOrder,'',SO_DATE,'','A',CURRENCY_CODE," &
                                " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                " '" & mP_User & "','1'," & " REMARKS,'E',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                                " '1','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                Else
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                " SELECT distinct CUST_CODE,SalesOrder,'',SO_DATE,'','A',CURRENCY_CODE," &
                                " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                " '" & mP_User & "','1'," & " REMARKS,'O',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                                " '1','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                End If
                                '10624527
                                Dim sbuilder As StringBuilder = New StringBuilder()
                                sbuilder.Append("INSERT INTO cust_ord_dtl (UNIT_CODE, Account_Code, Packing_Type, Cust_ref, Amendment_No, Item_Code, Rate, Order_Qty, Despatch_Qty, Active_Flag,")
                                sbuilder.Append(" Cust_Mtrl, Cust_DrgNo, Packing, Others, Excise_Duty, Cust_Drg_Desc, Tool_Cost, Authorized_flag, Ent_dt, Ent_UserId, Upd_dt, Upd_UserId,")
                                sbuilder.Append(" OpenSO, SalesTax_Type, PerValue, RevisionNo, TOOL_AMOR_FLAG, ShowInAuth, ADD_Excise_Duty, DISCOUNT_TYPE, DISCOUNT_VALUE,")
                                If gstrUNITID <> "MS1" Then      'INC1371722
                                    sbuilder.Append("HSNSACCODE, ")
                                End If
                                sbuilder.Append("CGSTTXRT_TYPE, SGSTTXRT_TYPE, IGSTTXRT_TYPE , COMPENSATION_CESS)")
                                sbuilder.Append(" SELECT '" & gstrUNITID & "',CUST_CODE,'PKT0',SalesOrder,'',item_code,exworks,Quantity,0,'A',")
                                sbuilder.Append("cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,packing_per,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid,")
                                sbuilder.Append(" getdate(), '" & mP_User & "', '1', VAT, '1', '0', '1', '1', AddEx, DISCOUNT_TYPE, DISCOUNT_VALUE ,")
                                If gstrUNITID <> "MS1" Then      'INC1371722
                                    sbuilder.Append("HSNSACCODE, ")
                                End If
                                sbuilder.Append("CGST_TAX, SGST_TAX, IGST_TAX, COMPCC_TAX  FROM SO_Upld_dtl")
                                sbuilder.Append(" WHERE UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'")

                                mP_Connection.Execute(sbuilder.ToString(), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    '10624527
                                Else
                                    If mblnSOUPLD_ForexCurrency = True Then
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                " SELECT distinct CUST_CODE,SalesOrder,'',SO_DATE,'','A',CURRENCY_CODE," &
                                " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                " '" & mP_User & "','1'," & " REMARKS,'E',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                                " '0','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                Else
                                    mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                                " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                                " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                                " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                                " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                                " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                                " SELECT distinct CUST_CODE,SalesOrder,'',SO_DATE,'','A',CURRENCY_CODE," &
                                " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                                " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                                " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                                " '" & mP_User & "','1'," & " REMARKS,'O',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                                " '0','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                                " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                End If

                                mP_Connection.Execute("INSERT INTO cust_ord_dtl (UNIT_CODE,Account_Code,Packing_Type,Cust_ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag," &
                                " Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId," &
                                " OpenSO,SalesTax_Type,PerValue,RevisionNo,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,DISCOUNT_TYPE,DISCOUNT_VALUE,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,IGSTTXRT_TYPE ,COMPENSATION_CESS)" &
                                " SELECT '" & gstrUNITID & "',CUST_CODE,'PKT0',SalesOrder,'',item_code,exworks,Quantity,0,'A'," &
                                "cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,packing_per,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid," &
                                " getdate(),'" & mP_User & "','0',VAT,'1','0','1','1',AddEx ,DISCOUNT_TYPE,DISCOUNT_VALUE ,HSNSACCODE,CGST_TAX,SGST_TAX,IGST_TAX,COMPCC_TAX  FROM SO_Upld_dtl" &
                                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If

                            'Issue id  10121169
                            'mP_Connection.Execute("INSERT INTO cust_ord_dtl (Account_Code,Cust_ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,SalesTax_Type,PerValue,RevisionNo,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty) SELECT CUST_CODE,SalesOrder,'',item_code,exworks,Quantity,0,'A'," & "cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,0.00,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid," & " getdate(),'" & mP_User & "','0','','1','0','1','1',AddEx FROM SO_Upld_dtl" & " WHERE DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            'Issue id  10121169 end
                        End If
                    Else
                        If mblnopensalesorder = True Then
                            If mblnSOUPLD_ForexCurrency = True Then
                                mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                    " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                    " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                    " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                    " SELECT distinct CUST_CODE,isnull(INTERNAL_SALESORDER_NO,''),'',SO_DATE,'','A',CURRENCY_CODE," &
                    " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                    " '" & mP_User & "','1'," & " REMARKS,'E',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                    " '1','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "'  FROM SO_Upld_dtl" &
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                     " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                    " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                    " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                    " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                    " SELECT distinct CUST_CODE,isnull(INTERNAL_SALESORDER_NO,''),'',SO_DATE,'','A',CURRENCY_CODE," &
                    " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                    " '" & mP_User & "','1'," & " REMARKS,'O',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                    " '1','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "'  FROM SO_Upld_dtl" &
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                     " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)


                            End If

                            mP_Connection.Execute("INSERT INTO cust_ord_dtl (UNIT_CODE,Account_Code,Packing_Type,Cust_ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag," &
                                   " Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId," &
                                    " OpenSO,SalesTax_Type,PerValue,RevisionNo,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,EXTERNAL_SALESORDER_NO,DISCOUNT_TYPE,DISCOUNT_VALUE,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,IGSTTXRT_TYPE ,COMPENSATION_CESS  )" &
                                    " SELECT '" & gstrUNITID & "',CUST_CODE,'PKT0',INTERNAL_SALESORDER_NO ,'',item_code,exworks,Quantity,0,'A'," &
                                    "cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,PACKING_PER,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid," &
                                    " getdate(),'" & mP_User & "','1',VAT,'1','0','1','1',AddEx ,salesorder ,DISCOUNT_TYPE,DISCOUNT_VALUE ,HSNSACCODE,CGST_TAX,SGST_TAX,IGST_TAX,COMPCC_TAX  FROM SO_Upld_dtl" &
                                    " WHERE DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                        Else
                            If mblnSOUPLD_ForexCurrency = True Then
                                mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                    " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                    " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                    " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                    " SELECT distinct CUST_CODE,isnull(INTERNAL_SALESORDER_NO,''),'',SO_DATE,'','A',CURRENCY_CODE," &
                    " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                    " '" & mP_User & "','1'," & " REMARKS,'E',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                    " '0','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "'  FROM SO_Upld_dtl" &
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                     " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute("insert into cust_ord_hdr (Account_Code,Cust_ref," &
                    " Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code," &
                    " Valid_Date, Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks," &
                     " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized," &
                    " Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type," &
                    " SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO," &
                    " AddCustSupp,PerValue,RevisionNo,surcharge_code,ECESS_Code,Consignee_Code, UNIT_CODE,CT2_Reqd_In_SO,ShipAddress_Code,ShipAddress_Desc)" &
                    " SELECT distinct CUST_CODE,isnull(INTERNAL_SALESORDER_NO,''),'',SO_DATE,'','A',CURRENCY_CODE," &
                    " VALID_TO,VALID_FROM,TERM_PAYMENT,Special_Remarks,Pay_Remarks," &
                    " Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks," &
                    " Octorai_Remarks,Mode_Despatch,Delivery,'" & mP_User & "','" & mP_User & "'," &
                    " '" & mP_User & "','1'," & " REMARKS,'O',VAT,GETDATE(),ENT_UID,GETDATE(),'" & mP_User & "'," &
                    " '0','',1,0,'',''," & " CONSIGNEE_CODE, '" & gstrUNITID & "', '" & mblnCT2Reqd & "','" & mstrshipaddcode & "','" & mstrshipadddesc & "' FROM SO_Upld_dtl" &
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND" &
                     " CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)


                            End If

                            'Issue id  10121169
                            'mP_Connection.Execute("INSERT INTO cust_ord_dtl (Account_Code,Cust_ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,SalesTax_Type,PerValue,RevisionNo,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty) SELECT CUST_CODE,SalesOrder,'',item_code,exworks,Quantity,0,'A'," & "cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,0.00,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid," & " getdate(),'" & mP_User & "','0','','1','0','1','1',AddEx FROM SO_Upld_dtl" & " WHERE DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute("INSERT INTO cust_ord_dtl (UNIT_CODE,Account_Code,Packing_Type,Cust_ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag," &
                                " Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId," &
                                " OpenSO,SalesTax_Type,PerValue,RevisionNo,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,EXTERNAL_SALESORDER_NO,DISCOUNT_TYPE,DISCOUNT_VALUE,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,IGSTTXRT_TYPE ,COMPENSATION_CESS  )" &
                                " SELECT '" & gstrUNITID & "',CUST_CODE,'PKT0',INTERNAL_SALESORDER_NO ,'',item_code,exworks,Quantity,0,'A'," &
                                "cast(CustSupply as Decimal(18,4))as CustSupply,PartNo,packing_per,0.00,EDper,PartName,ToolCost,'1',ent_dt,Ent_Uid," &
                                " getdate(),'" & mP_User & "','0',VAT,'1','0','1','1',AddEx ,salesorder ,DISCOUNT_TYPE,DISCOUNT_VALUE ,HSNSACCODE,CGST_TAX,SGST_TAX,IGST_TAX,COMPCC_TAX  FROM SO_Upld_dtl" &
                                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & txtDocNo.Text & "' AND CUST_CODE = '" & txtCustomerCode.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'Issue id  10121169 end

                        End If


                    End If
                    If mblnUPDATE_DMS_SO_SCH_UPLD = True Then
                        rsDlyMkt.GetResult("SELECT Account_Code,Trans_date,Item_code,Cust_Drgno," & " Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,RevisionNo," & " Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Consignee_code,filetype,Doc_No" & " FROM DAILYMKTSCHEDULE_TEMP Where Doc_No = '" & txtDocNo.Text & "' and UNIT_CODE = '" & gstrUNITID & "'")

                        If rsDlyMkt.GetNoRows > 0 Then
                            rsDlyMkt.MoveFirst()
                            While Not rsDlyMkt.EOFRecord
                                rsDup = New ClsResultSetDB
                                rsDup.GetResult("select  * from dailymktschedule" &
                                    " where UNIT_CODE = '" & gstrUNITID & "' AND" &
                                    " account_code = '" & rsDlyMkt.GetValue("account_code") & "'" &
                                    " and trans_date = '" & VB6.Format(rsDlyMkt.GetValue("trans_date"), "dd-MMM-yyyy") & "' " &
                                    " and item_code = '" & rsDlyMkt.GetValue("item_code") & "' " &
                                    " and consignee_code = '" & rsDlyMkt.GetValue("consignee_code") & "'")
                                If rsDup.GetNoRows > 0 Then
                                    rsDup.MoveFirst()
                                    While Not rsDup.EOFRecord
                                        mP_Connection.Execute("UPDATE DAILYMKTSCHEDULE" &
                                            " SET STATUS = 0, SCHEDULE_FLAG = 0" &
                                            " where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & rsDlyMkt.GetValue("account_code") & "'" &
                                            " and trans_date = '" & VB6.Format(rsDlyMkt.GetValue("trans_date"), "dd-MMM-yyyy") & "' " &
                                            " and item_code = '" & rsDlyMkt.GetValue("item_code") & "' " &
                                            " and consignee_code = '" & rsDlyMkt.GetValue("consignee_code") & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        rsDup.MoveNext()
                                    End While
                                End If
                                rsDup.ResultSetClose()
                                mP_Connection.Execute("INSERT INTO DAILYMKTSCHEDULE (Account_Code," &
                                " Trans_date,Item_code,Cust_Drgno,Schedule_Flag," &
                                " Schedule_Quantity,Despatch_Qty,Status,RevisionNo,Ent_dt,Ent_UserId," &
                                " Upd_dt,Upd_UserId,Consignee_code,filetype,Doc_No, UNIT_CODE)" &
                                " VALUES ('" & rsDlyMkt.GetValue("account_code") & "' ," &
                                " '" & VB6.Format(rsDlyMkt.GetValue("Trans_date"), "dd-MMM-yyyy") & "'," &
                                " '" & rsDlyMkt.GetValue("Item_code") & "'," &
                                " '" & rsDlyMkt.GetValue("Cust_Drgno") & "'," &
                                " '" & rsDlyMkt.GetValue("Schedule_Flag") & "' ," &
                                " '" & rsDlyMkt.GetValue("Schedule_Quantity") & "'," &
                                " '" & rsDlyMkt.GetValue("Despatch_Qty") & "'," &
                                " '" & rsDlyMkt.GetValue("Status") & "'," &
                                " '" & rsDlyMkt.GetValue("RevisionNo") & "'," &
                                " '" & VB6.Format(rsDlyMkt.GetValue("Ent_dt"), "dd-MMM-yyyy") & "','" & rsDlyMkt.GetValue("Ent_UserId") & "'," &
                                " '" & VB6.Format(rsDlyMkt.GetValue("Upd_dt"), "dd-MMM-yyyy") & "','" & rsDlyMkt.GetValue("Upd_UserId") & "'," &
                                " '" & rsDlyMkt.GetValue("Consignee_code") & "'," & " '" & rsDlyMkt.GetValue("filetype") & "'," &
                                " '" & rsDlyMkt.GetValue("Doc_No") & "', '" & gstrUNITID & "' )", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                rsDlyMkt.MoveNext()
                            End While
                        End If
                    End If
                    rsDlyMkt.ResultSetClose()
                    mP_Connection.Execute("UPDATE SO_UPLD_HDR SET CANAUTH = 0 WHERE DOC_NO = '" & Me.txtDocNo.Text & "' and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.CommitTrans()
                    MsgBox("Data Authorized Successfully.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                    cmdAuthorize.Enabled(0) = False
                    CmdUpdSchedule.Enabled = False

                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH
                Call RefreshForm()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
Errorhandler:
        mP_Connection.RollbackTrans()
        If Err.Number = -2147217900 Then
            MsgBox("Duplicate Entry.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelpAddExcise_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpAddExcise.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate WHERE UNIT_CODE = '" & gstrUNITID & "' ", "Additional Excise Duty", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtAddExcise.Text = strHelp(0)
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelpDocNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpDocNo.Click
        On Error GoTo ErrHandler
        Dim StrSql As String
        Dim docno As Integer
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        Call RefreshForm()
        ' StrSql = "SELECT distinct cast(H.DOC_NO as varchar(18)) as doc_no,M.CUSTOMER_CODE,M.Cust_Name,D.VALID_FROM, D.VALID_TO" & _
        '" FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H ON D.UNIT_CODE = H.UNIT_CODE AND D.DOC_NO = H.DOC_NO" & _
        '" AND H.DOC_NO NOT IN (SELECT DISTINCT ISNULL(DOC_NO,0)" & _
        '" FROM dailymkTschedule WHERE UNIT_CODE = '" & gstrUNITID & "')" & _
        '" INNER JOIN CUSTOMER_MST M ON" & _
        '" M.UNIT_CODE = H.UNIT_CODE AND H.UNIT_CODE = '" & gstrUNITID & "'" & _
        '" AND M.CUSTOMER_CODE = H.CUST_CODE AND ISNULL(H.CANAUTH,1) <> 0 and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date))"
        StrSql = "SELECT TOP 1 1  FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H ON D.UNIT_CODE = H.UNIT_CODE AND D.DOC_NO = H.DOC_NO " & _
                " INNER JOIN CUSTOMER_MST M ON      M.UNIT_CODE = H.UNIT_CODE " & _
                " AND M.CUSTOMER_CODE = H.CUST_CODE AND ISNULL(H.CANAUTH,1) <> 0" & _
                " and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date))" & _
                " where  H.UNIT_CODE = '" & gstrUNITID & "' AND D.QUANTITY='0' "

        If DataExist(StrSql) = True Then
            mblnopensalesorder = True
            StrSql = "SELECT distinct cast(H.DOC_NO as varchar(18)) as doc_no,M.CUSTOMER_CODE,M.Cust_Name,D.VALID_FROM, D.VALID_TO" & _
                    " FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H ON D.UNIT_CODE = H.UNIT_CODE AND D.DOC_NO = H.DOC_NO " & _
                    " INNER JOIN CUSTOMER_MST M ON      M.UNIT_CODE = H.UNIT_CODE" & _
                    " AND M.CUSTOMER_CODE = H.CUST_CODE AND ISNULL(H.CANAUTH,1) <> 0" & _
                    " and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date))" & _
                    " where  H.UNIT_CODE = '" & gstrUNITID & "'"
        Else
            mblnopensalesorder = False
            '10727228 
            StrSql = "SELECT distinct cast(H.DOC_NO as varchar(18)) as doc_no,M.CUSTOMER_CODE,M.Cust_Name,D.VALID_FROM, D.VALID_TO" & _
                     " FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H ON D.UNIT_CODE = H.UNIT_CODE AND D.DOC_NO = H.DOC_NO " & _
                    " INNER JOIN CUSTOMER_MST M ON      M.UNIT_CODE = H.UNIT_CODE" & _
                    " AND M.CUSTOMER_CODE = H.CUST_CODE AND ISNULL(H.CANAUTH,1) <> 0" & _
                    " and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date))" & _
                    " and NOT EXISTS (SELECT DISTINCT ISNULL(DOC_NO,0)" & _
                    " FROM dailymkTschedule WHERE UNIT_CODE = h.UNIT_CODE and DOC_NO = h.Doc_no and  FILETYPE ='SOUPLD' )" & _
                    " where  H.UNIT_CODE = '" & gstrUNITID & "'"
        End If
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "HELP", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) = "" Then
                    Exit Sub
                End If
                Me.txtDocNo.Text = strHelp(0)
                docno = CInt(txtDocNo.Text)
                Call txtDocNo_Validating(txtDocNo, New System.ComponentModel.CancelEventArgs(True))
                '        Call FetchSORecords(docno)
                cmdAuthorize.Enabled(0) = True
                cmdAuthorize.Enabled(1) = True
                cmdAuthorize.Enabled(2) = False
                cmdAuthorize.Enabled(3) = True
                mbln_View = True
                Ctrl_EnabledDisable()
                cmdSODetails.Enabled = False
            Else
                MsgBox(" No record available!", MsgBoxStyle.Information, ResolveResString(100))
                mbln_View = False
                cmdSODetails.Enabled = True
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelpEcess_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpEcess.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc" & _
            " from Gen_TaxRate WHERE UNIT_CODE = '" & gstrUNITID & "' ", "Help", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtECESS.Text = strHelp(0)
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelpSHECess_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpSHECess.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate WHERE UNIT_CODE = '" & gstrUNITID & "' ", "Help", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtSECESS.Text = strHelp(0)
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPlantHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPlantHelp.Click
        On Error GoTo ErrHandler
        Dim sql As String
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        sql = "select distinct C.wh_code ,W.WH_Description" & _
           " from custwarehouse_mst C, warehouse_mst W" & _
           " WHERE C.UNIT_CODE = W.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'" & _
           " AND customer_code = '" & txtCustomerCode.Text & "'" & _
           " and c.wh_code = w.wh_code"
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, sql, "Help", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtPlantCode.Text = strHelp(0)
                    Me.lblPlantName.Text = strHelp(1)
                    Me.txtPlantCode.Focus()
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdUpdSchedule_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdUpdSchedule.Click
        Dim varDayL, varDayF As Short
        Dim varDate As Object
        Dim varMonthyear As String
        Dim VarSchQty As Integer
        Dim rsWkngDays As New ClsResultSetDB
        Dim Row, Col As Short
        Dim varPARTNO As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varMonthyear = Month(dtpSOdate.Value) & "/" & Year(dtpSOdate.Value)
        Dim rsGetRows As New ClsResultSetDB
        If Val(txtDiff.Text) = 0 Then
            With SSPurOrd
                .Col = ENUM_Grid.InternalPartNo
                varItemCode = Nothing
                varPARTNO = Nothing
                .GetText(ENUM_Grid.InternalPartNo, .ActiveRow, varItemCode)
                .GetText(ENUM_Grid.CustPartNo, .ActiveRow, varPARTNO)
            End With
            With SpdSOSch
                If SSPurOrd.MaxRows > 0 Then
                    For Row = 1 To .MaxRows
                        For Col = 1 To .MaxCols
                            .Row = Row
                            .Col = Col
                            If .Value <> "" Then
                                VarSchQty = .Value
                                If Row = 1 Then
                                    varDayF = 0
                                    varDayL = Col - 1
                                ElseIf Row = 2 Then
                                    varDayF = 1
                                    varDayL = Col - 1
                                ElseIf Row = 3 Then
                                    varDayF = 2
                                    varDayL = Col - 1
                                ElseIf Row = 4 Then
                                    varDayF = 3
                                    varDayL = Col - 1
                                End If
                                varDate = varDayF & varDayL & "/" & varMonthyear
                                rsGetRows.GetResult("SELECT * FROM DAILYMKTSCHEDULE_TEMP" & _
                                   " WHERE UNIT_CODE = '" & gstrUNITID & "' AND" & _
                                   " Account_Code = '" & txtCustomerCode.Text & "' and" & _
                                   " Trans_date = '" & VB6.Format(varDate, "dd-MMM-yyyy") & "' and" & _
                                   " Item_code = '" & varItemCode & "' " & " AND FILETYPE = 'SOUPLD' and DOC_NO = '" & txtDocNo.Text & "'")
                                If rsGetRows.GetNoRows > 0 Then
                                    mP_Connection.Execute("UPDATE DailyMktSchedule_Temp SET Schedule_Quantity = " & VarSchQty & "" & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Account_Code = '" & Trim(txtCustomerCode.Text) & "'" & " and Trans_date = '" & VB6.Format(varDate, "dd-MMM-yyyy") & "' and" & " Item_code = '" & varItemCode & "' AND FILETYPE = 'SOUPLD' " & " AND DOC_NO = '" & txtDocNo.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                Else
                                    mP_Connection.Execute("INSERT INTO DailyMktSchedule_Temp (Account_Code," & _
                                        " Trans_date,Item_code,Cust_Drgno,Schedule_Flag," & _
                                        " Schedule_Quantity,Despatch_Qty,Status,Ent_dt," & _
                                        " Ent_UserId,Upd_dt,Upd_UserId,Consignee_code,doc_no,filetype,UNIT_CODE)" & _
                                        " VALUES('" & txtCustomerCode.Text & "','" & VB6.Format(varDate, "dd-MMM-yyyy") & "' ," & _
                                        " '" & varItemCode & "','" & varPARTNO & "'," & _
                                        " 1," & VarSchQty & " ,0,0,GETDATE(),'" & mP_User & "'," & _
                                        " GETDATE(),'" & mP_User & "','" & txtCustomerCode.Text & "','" & Trim(txtDocNo.Text) & "','SOUPLD','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If
                        Next
                    Next
                    MsgBox("Daily Marketing Schedule Updated Against Item:" & varItemCode, MsgBoxStyle.Information, ResolveResString(100))
                End If
            End With
        Else
            MsgBox("Actual quantity and schedule quantity must be same.", MsgBoxStyle.Information, ResolveResString(100))
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPrint_Click()
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        frmExport.ShowDialog()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        If gblnCancelExport Then Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) : Exit Sub
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdVATHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdVATHelp.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc" & _
            " from Gen_TaxRate WHERE UNIT_CODE = '" & gstrUNITID & "'", "Help", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtVat.Text = strHelp(0)
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("UNDERCONSTRUCTION.HTM")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0058_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0058_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0058_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader_ClickEvent(ctlFormHeader, New System.EventArgs())
        End If
    End Sub
    Private Sub frmMKTTRN0058_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Call FitToClient(Me, frmMain, ctlFormHeader, cmdAuthorize, 500)
        Call FillLabelFromResFile(Me)
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.HeaderString())
        Call EnableControls(True, Me)
        Call RefreshForm()
        cmbSearch.Items.Add("Item Code")
        cmbSearch.Items.Add("CustPart Code")
        txtDocNo.MaxLength = 18
        Call FN_Spread_Settings()
        txtinternalSO.Enabled = False
        txtinternalSO.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

        If IsRecordExists("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE UPDATE_DMS_SO_SCH_UPLD = 0 and UNIT_CODE = '" & gstrUNITID & "'") Then
            mblnUPDATE_DMS_SO_SCH_UPLD = False
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0058_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Me.Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Option1_KeyPress(ByRef KeyAscii As Short)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Option2_KeyPress(ByRef KeyAscii As Short)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub OptSalesordPrint_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptSalesordPrint.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub OptSchPrint_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptSchPrint.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub SpdSOSch_Change1(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpdSOSch.Change
        Dim SCHQTY As Long
        On Error GoTo ErrHandler
        With SpdSOSch
            For e.row = 1 To .MaxRows
                For e.col = 1 To .MaxCols
                    .Row = e.row : .Col = e.col
                    If .Value = "" Then
                        SCHQTY = SCHQTY + 0
                    Else
                        SCHQTY = SCHQTY + .Value
                    End If
                Next
            Next
        End With
        txtSchQty.Text = SCHQTY
        txtDiff.Text = Val(txtQty.Text) - Val(txtSchQty.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpdSOSch_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpdSOSch.KeyPressEvent
        'Dim KeyAscii As Short = Asc(e.keyAscii)
        If e.keyAscii = 45 Or e.keyAscii = 43 Then
            e.keyAscii = 0
        End If
    End Sub
    Private Sub SSPurOrd_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SSPurOrd.ClickEvent
        On Error GoTo Errorhandler
        Dim varInternalPart As Object, varCustPart As Object
        Dim rsitemdesc As New ClsResultSetDB
        With SSPurOrd
            varInternalPart = Nothing
            varCustPart = Nothing
            .Row = e.row
            .GetText(ENUM_Grid.InternalPartNo, .Row, varInternalPart)
            .GetText(ENUM_Grid.CustPartNo, .Row, varCustPart)
        End With
        rsitemdesc.GetResult("SELECT ITEM_DESC,DRG_DESC FROM CUSTITEM_MST" & _
           " WHERE UNIT_CODE = '" & gstrUNITID & "' AND ITEM_CODE = '" & varInternalPart & "'" & _
           " AND CUST_DRGNO = '" & varCustPart & "' AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "'" & _
           " AND ACTIVE = 1")
        If rsitemdesc.GetNoRows > 0 Then
            lblItemDesc.Text = rsitemdesc.GetValue("ITEM_DESC")
            lblCustPartDesc.Text = rsitemdesc.GetValue("DRG_DESC")
        Else
            lblItemDesc.Text = ""
            lblCustPartDesc.Text = ""
        End If
        Call DisplayDailyMktSchedule(True, 0)
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SSPurOrd_LeaveCell1(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SSPurOrd.LeaveCell
        On Error GoTo Errorhandler
        If e.newRow >= 1 And e.newRow <= SSPurOrd.MaxRows Then
            Call DisplayDailyMktSchedule(False, e.newRow)
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        On Error GoTo Errorhandler
        LblCustomerName.Text = ""
        txtPlantCode.Text = ""
        lblPlantName.Text = ""
        SSPurOrd.MaxRows = 0
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdCustHelp_Click(cmdCustHelp, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim bolExist As Boolean
        If Len(txtCustomerCode.Text) > 0 Then
            bolExist = ValCustomerCode()
        End If
        If bolExist = True Then
            Call display()
        End If
        blnlinelevelcustomer = Find_Value("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")
        If blnlinelevelcustomer = True Then
            txtinternalSO.Text = Find_Value("SELECT TOP 1 isnull(INTERNAL_SALESORDER_NO,'')  INTERNAL_SALESORDER_NO FROM SO_UPLD_DTL WHERE DOC_NO='" & txtDocNo.Text.Trim & "' and UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")
        End If
        mblnSOUPLD_ForexCurrency = Find_Value("SELECT SOUPLD_FOREXCURRENCY FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")

        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtDocNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocNo.TextChanged
        On Error GoTo ErrHandler
        Flag = True
        If Len(Trim(txtDocNo.Text)) = 0 Then
            Call RefreshForm()
        End If
        Flag = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDocNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim docno As Integer
        Dim StrSql As String
        Dim rsDocNo As New ClsResultSetDB
        If Len(txtDocNo.Text) > 0 Then
            docno = CInt(txtDocNo.Text)
            Call RefreshForm()
            txtDocNo.Text = CStr(docno)
            StrSql = "SELECT TOP 1 1  FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H ON D.UNIT_CODE = H.UNIT_CODE AND D.DOC_NO = H.DOC_NO " & _
                " INNER JOIN CUSTOMER_MST M ON      M.UNIT_CODE = H.UNIT_CODE " & _
                " AND M.CUSTOMER_CODE = H.CUST_CODE AND ISNULL(H.CANAUTH,1) <> 0" & _
                " and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date))" & _
                " where  H.UNIT_CODE = '" & gstrUNITID & "' AND D.QUANTITY='0' AND H.DOC_NO='" & Me.txtDocNo.Text & "'"
            If DataExist(StrSql) = True Then
                mblnopensalesorder = True
                StrSql = "SELECT distinct H.DOC_NO,M.CUSTOMER_CODE,M.Cust_Name,D.VALID_FROM," & _
                    " D.VALID_TO FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H" & _
                    " ON D.UNIT_CODE = H.UNIT_CODE AND H.UNIT_CODE = '" & gstrUNITID & "'" & _
                    " AND D.DOC_NO = H.DOC_NO " & _
                    " INNER JOIN CUSTOMER_MST M ON" & _
                    " M.UNIT_CODE = H.UNIT_CODE AND M.CUSTOMER_CODE = H.CUST_CODE AND ISNULL(H.CANAUTH,1) <> 0 and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date)) "
            Else
                mblnopensalesorder = False
                '10727228 
                StrSql = "SELECT distinct H.DOC_NO,M.CUSTOMER_CODE,M.Cust_Name,D.VALID_FROM," & _
                    " D.VALID_TO FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H" & _
                    " ON D.UNIT_CODE = H.UNIT_CODE AND H.UNIT_CODE = '" & gstrUNITID & "'" & _
                    " AND D.DOC_NO = H.DOC_NO" & _
                    " AND H.DOC_NO NOT IN (SELECT DISTINCT ISNULL(DOC_NO,0)" & _
                    " FROM dailymkTschedule WHERE UNIT_CODE = '" & gstrUNITID & "' and  FILETYPE ='SOUPLD' )" & _
                    " AND H.DOC_NO = '" & docno & "'" & _
                    " INNER JOIN CUSTOMER_MST M ON" & _
                    " M.UNIT_CODE = H.UNIT_CODE AND M.CUSTOMER_CODE = H.CUST_CODE and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date)) "
            End If
            rsDocNo.GetResult(StrSql)
            If rsDocNo.GetNoRows > 0 Then
                Call FetchSORecords(docno)
                cmdAuthorize.Enabled(0) = True
                cmdAuthorize.Enabled(1) = True
                cmdAuthorize.Enabled(2) = False
                cmdAuthorize.Enabled(3) = True
                mbln_View = True
                Ctrl_EnabledDisable()
                cmdSODetails.Enabled = False
                Me.cmdAuthorize.Focus()
            Else
                MsgBox(" No record available!", MsgBoxStyle.Information, ResolveResString(100))
                RefreshForm()
                mbln_View = False
                cmdSODetails.Enabled = True
            End If
            If txtCustomerCode.Text.Length > 0 Then
                blnlinelevelcustomer = Find_Value("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")
                If blnlinelevelcustomer = True Then
                    txtinternalSO.Text = Find_Value("SELECT TOP 1 isnull(INTERNAL_SALESORDER_NO,'')  INTERNAL_SALESORDER_NO FROM SO_UPLD_DTL WHERE DOC_NO='" & txtDocNo.Text.Trim & "' and UNIT_CODE='" & gstrUNITID & "'")
                End If
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtECESS_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtECESS.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        'if KeyAscii
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtExciseDuty_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtExciseDuty.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdExDutyHelp_Click(cmdExDutyHelp, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAddExcise_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAddExcise.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdHelpAddExcise_Click(cmdHelpAddExcise, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtExciseDuty_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtExciseDuty.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSearch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSECESS_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSECESS.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtVat_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtVat.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdVATHelp_Click(cmdVATHelp, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtEcess_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtECESS.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdHelpEcess_Click(cmdHelpEcess, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSECESS_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSECESS.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdHelpSHECess_Click(cmdHelpSHECess, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtPlantCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPlantCode.TextChanged
        On Error GoTo Errorhandler
        If Len(txtPlantCode.Text) <= 0 Then
            lblPlantName.Text = ""
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtplantcode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPlantCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdPlantHelp_Click(cmdPlantHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtplantcode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPlantCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtplantcode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtPlantCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo Errorhandler
        Dim rsobject As New ClsResultSetDB
        If Len(Trim(Me.txtPlantCode.Text)) > 0 Then
            If Not CheckRecord("select distinct C.wh_code as PlantCode,W.WH_Description as PlantName" & _
                " from custwarehouse_mst C, warehouse_mst W" & _
                " where C.UNIT_CODE = W.UNIT_CODE AND C.UNIT_CODE = '" & gstrUNITID & "'" & _
                " AND customer_code = '" & txtCustomerCode.Text & "'" & _
                " and c.wh_code = w.wh_code AND W.WH_CODE = '" & txtPlantCode.Text & "' ") Then
                MsgBox(" Invalid Plant Code", MsgBoxStyle.Information, ResolveResString(100))
                Me.txtPlantCode.Text = "" : Me.txtPlantCode.Focus()
                Cancel = True
            End If
        End If
        If Cancel = True Then
            Me.txtPlantCode.Focus()
        Else
            Call rsobject.GetResult("Select WH_Description from warehouse_mst" & _
                " where UNIT_CODE = '" & gstrUNITID & "' AND wh_code = '" & Trim(Me.txtPlantCode.Text) & "'")
            If rsobject.RowCount > 0 Then
                Me.lblPlantName.Text = rsobject.GetValue("WH_Description")
            Else
                Me.lblPlantName.Text = ""
            End If
        End If
        rsobject.ResultSetClose()
        rsobject = Nothing
        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function CheckRecord(ByRef StrSql As String) As Boolean
        On Error GoTo Errorhandler
        Dim rsobject As New ClsResultSetDB
        rsobject.GetResult(StrSql)
        If rsobject.RowCount > 0 Then
            CheckRecord = True
        Else
            CheckRecord = False
        End If
        rsobject.ResultSetClose()
        rsobject = Nothing
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtFileName_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFileName.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdFileHelp_Click(cmdFileHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtFileName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFileName.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Function RefreshForm() As Object
        On Error GoTo Errorhandler
        Dim Row, Col As Short
        Dim rsDtpDateTo As New ClsResultSetDB
        Dim StrSql As String
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        txtCustomerCode.Text = ""
        txtCustomerCode.Enabled = False
        txtPlantCode.Text = ""
        txtPlantCode.Enabled = False
        txtFileName.Text = ""
        txtFileName.Enabled = False
        SSPurOrd.MaxRows = 0
        txtExciseDuty.Text = ""
        txtAddExcise.Text = ""
        txtVat.Text = ""
        txtECESS.Text = ""
        txtSECESS.Text = ""
        txtSchQty.Text = ""
        txtQty.Text = ""
        txtDiff.Text = ""
        If Flag = False Then
            txtDocNo.Text = ""
        End If
        txtRemarks.Enabled = False
        txtRemarks.Text = ""
        cmdSODetails.Enabled = False
        cmdCustHelp.Enabled = False
        cmdPlantHelp.Enabled = False
        cmdFileHelp.Enabled = False
        dtpDateFrom.Enabled = False
        dtpDateFrom.Format = DateTimePickerFormat.Custom
        dtpDateFrom.CustomFormat = gstrDateFormat
        dtpDateFrom.Value = GetServerDate()
        txtCreditTermId.Text = ""
        txtCredirDesc.Text = ""
        'changes done against issue ID 10709130 by Abhinav
        If gstrUNITID <> "MPU" Then
            dtpDateTo.Enabled = False
            dtpDateTo.Format = DateTimePickerFormat.Custom
            dtpDateTo.CustomFormat = gstrDateFormat
            dtpDateTo.Value = GetServerDate()
        Else
            StrSql = "Select Top 1 Fin_End_date from Financial_Year_Tb where UNIT_CODE= '" & gstrUNITID & "' and GETDATE() between Fin_Start_date and Fin_End_date"
            If IsRecordExists(StrSql) = True Then
                rsDtpDateTo.GetResult(StrSql)
                dtpDateTo.Text = VB6.Format(rsDtpDateTo.GetValue("Fin_End_date"), gstrDateFormat)
            Else
                dtpDateTo.Text = GetServerDate()
            End If
        End If
        rsDtpDateTo.ResultSetClose()
        rsDtpDateTo = Nothing
        'exterminates here (10709130)
        dtpSOdate.Enabled = False
        dtpSOdate.Format = DateTimePickerFormat.Custom
        dtpSOdate.CustomFormat = gstrDateFormat
        dtpSOdate.Value = GetServerDate()
        lblCustPartDesc.Text = ""
        lblItemDesc.Text = ""
        txtSearch.Text = ""
        txtinternalSO.Text = ""
        Call FN_Spread_Settings()
        For Row = 1 To 4
            For Col = 1 To 10
                SpdSOSch.SetText(Col, Row, "")
            Next
        Next
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
Errorhandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub CboSearch(ByRef SSPurOrd As AxFPSpreadADO.AxfpSpread, ByRef txtSearch As System.Windows.Forms.TextBox, Optional ByRef cmbSearch As System.Windows.Forms.ComboBox = Nothing)
        On Error GoTo Err_Handler
        Dim intCount As Short
        Dim varSearchText As Object
        Dim strSearchText As String
        If Len(Trim(txtSearch.Text)) = 0 Then Exit Sub
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With SSPurOrd
            For intCount = 1 To .MaxRows
                varSearchText = Nothing
                Call .GetText(IIf(cmbSearch.SelectedIndex = 0, ENUM_Grid.InternalPartNo, ENUM_Grid.CustPartNo), intCount, varSearchText)
                strSearchText = Mid(CStr(varSearchText), 1, Len(Trim(txtSearch.Text)))
                If Trim(UCase(strSearchText)) = Trim(UCase(txtSearch.Text)) Then
                    .Row = intCount : .Col = IIf(cmbSearch.SelectedIndex = 0, ENUM_Grid.InternalPartNo, ENUM_Grid.CustPartNo) : .TopRow = .Row
                    .Col = ENUM_Grid.InternalPartNo
                    .Col2 = ENUM_Grid.CustPartNo
                    .Row2 = intCount
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                    .Font = VB6.FontChangeBold(.Font, True)
                    .Col = ENUM_Grid.CustPartNo
                    .Font = VB6.FontChangeBold(.Font, True)
                    .Col = ENUM_Grid.InternalPartNo
                    .Col2 = ENUM_Grid.CustPartNo
                    .Row2 = intCount
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                    Exit For
                Else
                    .Row = intCount : .Col = 1
                    .Col = ENUM_Grid.InternalPartNo
                    .Col2 = ENUM_Grid.CustPartNo
                    .Row2 = intCount
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                    .Font = VB6.FontChangeBold(.Font, False)
                    .Col = ENUM_Grid.CustPartNo
                    .Font = VB6.FontChangeBold(.Font, False)
                    .Col = ENUM_Grid.InternalPartNo
                    .Col2 = ENUM_Grid.CustPartNo
                    .Row2 = intCount
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End If
            Next intCount
            For intCount = intCount + 1 To .MaxRows
                .Row = intCount : .Col = 1
                .Col = ENUM_Grid.InternalPartNo
                .Col2 = ENUM_Grid.CustPartNo
                .Row2 = intCount
                .BlockMode = True
                .Lock = False
                .BlockMode = False
                .Font = VB6.FontChangeBold(.Font, False)
                .Col = ENUM_Grid.CustPartNo
                .Font = VB6.FontChangeBold(.Font, False)
                .Col = ENUM_Grid.InternalPartNo
                .Col2 = ENUM_Grid.CustPartNo
                .Row2 = intCount
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            Next
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
Err_Handler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearch.TextChanged
        On Error GoTo ErrHandler
        Call CboSearch(SSPurOrd, txtSearch, cmbSearch)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtVat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVat.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Function ValCustomerCode() As Boolean ' Checks for Customer Code
        On Error GoTo ErrHandler
        Dim ms As String
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        Call mobjEmpDll.CRecordset.OpenRecordset("select * from customer_mst" & _
            " where UNIT_CODE = '" & gstrUNITID & "' AND customer_code = '" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        mobjEmpDll.CRecordset.Filter_Renamed = "customer_code='" & Trim(txtCustomerCode.Text) & "'"
        If mobjEmpDll.CRecordset.Recordcount > 0 Then ValCustomerCode = True Else ValCustomerCode = False
        mobjEmpDll.CConnection.CloseConnection()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    Public Function display() As Object ' Used to display the record.
        On Error GoTo ErrHandler
        Dim strdetails As String
        Dim StrSql As String
        Dim strcheck As String
        Dim strmilk As Boolean
        Dim rsLocation As New ClsResultSetDB
        If Me.txtCustomerCode.Text <> "" Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mobjEmpDll.CRecordset.OpenRecordset("select * from customer_mst" & _
                " where UNIT_CODE = '" & gstrUNITID & "' AND Customer_code = '" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rsLocation.GetResult("select * from customer_mst where UNIT_CODE = '" & gstrUNITID & "' AND" & _
               " Customer_code = '" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
            If Not mobjEmpDll.CRecordset.EOF_Renamed Then
                mobjEmpDll.CRecordset.MoveFirst()
                txtCustomerCode.Text = mobjEmpDll.CRecordset.GetFieldValue("customer_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                LblCustomerName.Text = mobjEmpDll.CRecordset.GetFieldValue("cust_name", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
            End If
        End If
        rsLocation.ResultSetClose()

        Return Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    Function FetchSORecords(ByRef docno As Integer) As String
        Dim StrSql As String
        Dim Row, Col As Integer
        Dim rsRecords As ClsResultSetDB
        Dim rsWkngDays As New ClsResultSetDB
        Dim varDespatchQty As Object = Nothing
        Dim rsPrevQty As ADODB.Recordset
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call FN_Spread_Settings()
        Row = 1
        StrSql = "SELECT D.*,M.CUST_NAME,h.plant_c,h.file_location," & _
           " h.remarks AS REMARKS_HDR ,W.WH_Description,packing_per" & _
           " FROM dbo.SO_Upld_hdr h INNER JOIN dbo.SO_Upld_dtl D" & _
           " ON H.UNIT_CODE = D.UNIT_CODE AND" & _
           " h.Doc_no = D.Doc_no AND" & _
           " h.cust_code = D.CUST_CODE" & _
           " left outer  JOIN CUSTOMER_MST M ON M.UNIT_CODE = H.UNIT_CODE AND" & _
           " M.CUSTOMER_CODE = H.CUST_CODE" & _
           " left outer JOIN WAREHOUSE_MST W ON W.UNIT_CODE = H.UNIT_CODE AND" & _
           " W.WH_CODE = H.PLANT_C " & _
           " where h.DOC_NO = '" & txtDocNo.Text & "' AND H.UNIT_CODE = '" & gstrUNITID & "' order by SalesOrder"
        rsRecords = New ClsResultSetDB
        rsRecords.GetResult(StrSql)
        If rsRecords.RowCount > 0 Then
            rsRecords.MoveFirst()
            txtCustomerCode.Text = rsRecords.GetValue("cust_code")
            txtPlantCode.Text = rsRecords.GetValue("plant_c")
            txtFileName.Text = rsRecords.GetValue("file_location")
            dtpDateFrom.Value = rsRecords.GetValue("valid_from")
            dtpDateTo.Value = rsRecords.GetValue("valid_to")
            dtpSOdate.Value = rsRecords.GetValue("so_date")
            LblCustomerName.Text = rsRecords.GetValue("CUST_NAME")
            lblPlantName.Text = rsRecords.GetValue("WH_DESCRIPTION")
            While Not rsRecords.EOFRecord
                With SSPurOrd
                    'Row = .maxRows
                    SSPurOrd.MaxRows = SSPurOrd.MaxRows + 1
                    .SetText(ENUM_Grid.salesorder, .MaxRows, rsRecords.GetValue("SalesOrder"))
                    .SetText(ENUM_Grid.CustPartNo, .MaxRows, rsRecords.GetValue("PartNo"))
                    .SetText(ENUM_Grid.InternalPartNo, .MaxRows, rsRecords.GetValue("item_code"))
                    .SetText(ENUM_Grid.ItemDesc, .MaxRows, rsRecords.GetValue("PartName"))
                    .SetText(ENUM_Grid.CustPartDesc, .MaxRows, rsRecords.GetValue("SalesOrder"))
                    .SetText(ENUM_Grid.Qty, .MaxRows, rsRecords.GetValue("Quantity"))
                    .SetText(ENUM_Grid.ExWorks, .MaxRows, rsRecords.GetValue("ExWorks"))
                    .SetText(ENUM_Grid.custsupply, .MaxRows, rsRecords.GetValue("CustSupply"))
                    .SetText(ENUM_Grid.ToolCost, .MaxRows, rsRecords.GetValue("ToolCost"))
                    .SetText(ENUM_Grid.EDPer, .MaxRows, rsRecords.GetValue("EDPer"))
                    .SetText(ENUM_Grid.AddEx, .MaxRows, rsRecords.GetValue("AddEx"))
                    .SetText(ENUM_Grid.Cess, .MaxRows, rsRecords.GetValue("Cess"))
                    .SetText(ENUM_Grid.SHECess, .MaxRows, rsRecords.GetValue("SHECess"))
                    .SetText(ENUM_Grid.VAT, .MaxRows, rsRecords.GetValue("VAT"))
                    .SetText(ENUM_Grid.Remarks, .MaxRows, rsRecords.GetValue("Remarks"))
                    '.Col = ENUM_Grid.HSNSACCODE
                    If gstrUNITID <> "MS1" Then     'INC1371722
                        .SetText(ENUM_Grid.HSNSACCODE, .MaxRows, Convert.ToString(rsRecords.GetValue("HSNSACCODE")))
                    End If
                    .SetText(ENUM_Grid.CGST_TAX, .MaxRows, rsRecords.GetValue("CGST_TAX"))
                    .SetText(ENUM_Grid.SGST_TAX, .MaxRows, rsRecords.GetValue("SGST_TAX"))
                    .SetText(ENUM_Grid.IGST_TAX, .MaxRows, rsRecords.GetValue("IGST_TAX"))
                    .SetText(ENUM_Grid.PACKING_PER, .MaxRows, rsRecords.GetValue("packing_per"))

                    '10624527
                    If rsRecords.GetValue("DISCOUNT_TYPE") <> "" Then
                        .SetText(ENUM_Grid.discount_type, .MaxRows, rsRecords.GetValue("DISCOUNT_TYPE"))
                    End If
                    If rsRecords.GetValue("DISCOUNT_VALUE") <> "" Then
                        .SetText(ENUM_Grid.discount_value, .MaxRows, rsRecords.GetValue("DISCOUNT_VALUE"))
                    End If
                    '10624527
                    If rsRecords.GetValue("PrevQty") <> "" Then
                        .SetText(ENUM_Grid.PrevQty, .MaxRows, CInt(rsRecords.GetValue("PrevQty")))
                    End If
                    If rsRecords.GetValue("DespatchQty") <> "" Then
                        .SetText(ENUM_Grid.DespatchQty, .MaxRows, CInt(rsRecords.GetValue("DespatchQty")))
                    End If
                End With
                rsRecords.MoveNext()
            End While
            cmdAuthorize.Enabled(0) = True
            rsWkngDays.GetResult("select distinct work_flg,dt from" & _
                " calendar_mfg_mst where UNIT_CODE = '" & gstrUNITID & "' AND" & _
                " month (dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "'" & _
                " and YEAR (dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' order by dt")
            If rsWkngDays.GetNoRows > 0 Then
                rsWkngDays.MoveFirst()
                For Row = 1 To 4
                    For Col = 1 To 10
                        If Row = 1 And Col = 1 Then
                            SpdSOSch.Col = Col
                            SpdSOSch.Row = Row
                            SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            GoTo skip
                        End If
                        SpdSOSch.Col = Col
                        SpdSOSch.Row = Row
                        If rsWkngDays.EOFRecord = False Then
                            If rsWkngDays.GetValue("work_flg") = True Then
                                SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Else
                                SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                            End If
                            rsWkngDays.MoveNext()
                        Else
                            SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        End If
skip:
                    Next
                Next
            Else
                For Row = 1 To 4
                    For Col = 1 To 10
                        SpdSOSch.Col = Col
                        SpdSOSch.Row = Row
                        SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    Next
                Next
            End If
            SpdSOSch.Col = 1
            SpdSOSch.Col2 = 1
            SpdSOSch.Row = 4
            SpdSOSch.Row2 = 4
            If (SpdSOSch.BackColor) = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) Then
                SpdSOSch.BlockMode = True
                SpdSOSch.Lock = False
                SpdSOSch.CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                SpdSOSch.BlockMode = False
            End If
            For Row = 1 To 4
                For Col = 1 To 10
                    SpdSOSch.SetText(Col, Row, "")
                Next
            Next
            rsWkngDays.ResultSetClose()
            Call DisplayDailyMktSchedule(False, 1)
            If mblnUPDATE_DMS_SO_SCH_UPLD = False Then
                CmdUpdSchedule.Enabled = False
            Else
                CmdUpdSchedule.Enabled = True
            End If
            'added by priti 
            txtCreditTermId.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar(" select credit_days from customer_mst where unit_code ='" + gstrUNITID + "' and customer_code='" + txtCustomerCode.Text.Trim() + "'"))
            txtCredirDesc.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar(" select crtrm_desc from Gen_CreditTrmMaster where unit_code ='" + gstrUNITID + "' and CrTrm_TermId='" + txtCreditTermId.Text.Trim() + "'"))
        Else
            FetchSORecords = "No Record Found"
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Return Nothing
        Exit Function
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub Ctrl_EnabledDisable()
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        If mbln_View = True Then
            txtCustomerCode.Enabled = False
            txtPlantCode.Enabled = False
            txtFileName.Enabled = False
            txtRemarks.Enabled = False
            dtpDateFrom.Enabled = False
            dtpDateTo.Enabled = False
            dtpSOdate.Enabled = False
            cmdCustHelp.Enabled = False
            cmdPlantHelp.Enabled = False
            cmdFileHelp.Enabled = False
        ElseIf mbln_View = False Then
            Call RefreshForm()
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDocNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdHelpDocNo_Click(cmdHelpDocNo, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDocNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Dim docno As Integer
        Dim StrSql As String
        Dim rsDocNo As New ClsResultSetDB
        If KeyAscii = 8 Then
            GoTo EventExitSub
        End If
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii <> 13 Then
                KeyAscii = 0
            End If
        End If
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Function Find_Value(ByRef strField As String) As String
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
End Class