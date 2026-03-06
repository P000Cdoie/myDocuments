Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data.SqlClient
Friend Class frmMKTTRN0057
	Inherits System.Windows.Forms.Form
	'****************************************************
	'Copyright (c)  -  MIND
	'Name of module -  FRMMKTTRN0057.frm
	'Created By     -  Shubhra Verma
    'Created On     -  12 Oct 2007
    'Issue ID       -  21110
    'description    -  Sales Order And Schedule Uploading
    'Revised By     -  Shubhra Verma
    'Revised On     -  20 Jan 2009 to 21 Jan 2009
    'Issue ID       -  eMpro-20090120-26215
    'Revised History-  If Schedule is for current Month, then system should divide that required qty in remaining
    '                  working days of that month. 
    '                  if some Qty is already despatched for the item, then system should use New Schedule Qty - 
    '                  Despatched Qty as New Schedule Qty
    'Revised By     -  Shubhra Verma
    'Revised On     -  07 Jun 2010
    'Issue ID       -  eMpro-20100607-48114
    'Descri[ption   -  In So and Schedule uploading form, system should consider despatches 
    '                  between valid from and current date
    'Modified By JY on 17-May-2011
    '   Modified to support MultiUnit functionality
    ''**********************************************************************************
    'Revised By         :       SAURAV KUMAR    
    'Issue Id           :       10114399
    'Revised date       :       14 JULY 2011
    'Reason             :       Conversion from string "" to type double was not valid while uploading the excel file

    'Modified By        :       SAURAV KUMAR
    'Modified on        :       02 Aug 2011
    'Issue Id           :
    'Reason             :       String empty value not convertible into decimal

    'Modified By        :       Prashant Rajpal
    'Modified on        :       01 Sep 2011
    'Issue Id           :       10133711 
    'Reason             :       USER IS NOT SUPPOSED TO CHANGED SO DATE , VALIDATION CODE WAS INCORRECT , NOW IT'S OK,
    '**********************************************************************************
    'MODIFIED BY VIRENDRA GUPTA 0N 09 NOV 2011 FOR CHANGE MANAGEMENT
    '********************************************************************************************************
    'Modified By        :       Prashant Rajpal
    'Modified on        :       09 jan 2012
    'Issue Id           :       10178100  
    'Reason             :       shop code and UOM Added in So And schedule uploading (format is also changed )
    '**********************************************************************************
    'MODIFIED BY NITIN MEHTA 0N 31 JAN 2012 FOR CHANGE MANAGEMENT
    '**********************************************************************************
    'Modified By Shubhra on 31 Mar 2012 to increase the length of Doc No
    '**********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10229989
    'Revision Date   : 10-aug -2013- 31-aug -2013
    'History         : Multiple So Functionlity 
    '****************************************************************************************
    'Revised By      : Shubhra Verma
    'Issue ID        : 10530629
    'Revision Date   : 04 Feb 2014
    'History         : Check Added for Tariff Code while creating SO.
    '****************************************************************************************
    'Revised By      : Prashant Rajpal
    'Issue ID        : 10624527
    'Revision Date   : 18-SEP-2014
    'History         : Check TOOL COST AND DISCOUNT VALUE IN SO AND SCHEDULE UPLOADING FORM 
    '****************************************************************************************
    'Revised By      : Parveen Kumar
    'Issue ID        : 10725400
    'Revision Date   : 16-DEC-2014
    'History         : SO uploading valid to date 
    '****************************************************************************************
    'Revised By      : Abhinav Kumar
    'Issue ID        : 10736222
    'Revision Date   : 13 Jan 2015
    'History         : Changes done for CT2 ARE 3 functionality
    '****************************************************************************************
    'Created By     : Parveen Kumar
    'Created On     : 13 FEB 2015
    'Description    : eMPro Vehicle BOM
    'Issue ID       : 10737738 
    '-----------------------------------------------------------------------------------------
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10797956  
    'DESCRIPTION        :   EMPRO-CHANGES IN CT2 ARE-3 FUNCTIONALITY 
    'REVISION DATE      :   30 APR 2015
    '***************************************************************************************
    'REVISED BY         :   Parveen Kumar
    'ISSUE ID           :   10808160
    'DESCRIPTION        :   eMPro-New functionality of EOP
    'REVISION DATE      :   23 JUN 2015
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
    Dim bool_value_change As Boolean = False
    Dim blnlinelevelcustomer As Boolean = False
    Dim mblnDUPLICATE_SOALLOWED As Boolean = False
    Dim SchUpdFlag As Boolean = False   '10737738
    Dim mblnUPDATE_DMS_SO_SCH_UPLD As Boolean = True
    Dim blnSOUPLD_NotRemovehypen As Boolean = False
    Dim mblnSOUPLD_ForexCurrency As Boolean = False
    Dim mblnshippingcodemandatory As Boolean = False


    Private Enum ENUM_Grid
        ' check = 1
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
        Packing_Per
        EDPer
        AddEx
        Cess
        SHECess
        VAT
        ShopCode
        UOM
        'plantcode
        '10624527
        discount_type
        discount_value
        '10624527
        'GST CHANGES
        CGST_TAX
        SGST_TAX
        IGST_TAX
        COMP_ECC
        HSN_CODE
        'GST CHANGES
        Remarks

    End Enum
    Private Sub FN_Spread_Settings()
        Dim Col As Short
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With SSPurOrd
            'If gblnGSTUnit = False Then
            .MaxCols = ENUM_Grid.Remarks
            'Else
            '.MaxCols = ENUM_Grid.HSN_CODE
            'End If
            '.MaxCols = ENUM_Grid.Remarks
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
            If gblnGSTUnit = False Then
                .Col = ENUM_Grid.custsupply : .Text = "Cust Supply"
                .Col = ENUM_Grid.ToolCost : .Text = "Tool Cost"
                .Col = ENUM_Grid.Packing_Per : .Text = "Packing Per unit"
                .Col = ENUM_Grid.EDPer : .Text = "E.D.%"
                .Col = ENUM_Grid.AddEx : .Text = "AddE.D.%"
                .Col = ENUM_Grid.SHECess : .Text = "SHECess%"
                .Col = ENUM_Grid.Cess : .Text = "Cess%"
                .Col = ENUM_Grid.VAT : .Text = "VAT%"
            Else
                .Col = ENUM_Grid.custsupply : .Text = "Cust Supply" : .ColHidden = True
                If UCase(Trim(gstrUNITID)) = "MST" Or UCase(Trim(gstrUNITID)) = "MSB" Or UCase(Trim(gstrUNITID)) = "SDA" Then
                    .Col = ENUM_Grid.ToolCost : .Text = "Tool Cost" : .ColHidden = False
                Else
                    .Col = ENUM_Grid.ToolCost : .Text = "Tool Cost" : .ColHidden = True
                End If
                .Col = ENUM_Grid.EDPer : .Text = "E.D.%" : .ColHidden = True
                .Col = ENUM_Grid.AddEx : .Text = "AddE.D.%" : .ColHidden = True
                .Col = ENUM_Grid.SHECess : .Text = "SHECess%" : .ColHidden = True
                .Col = ENUM_Grid.Cess : .Text = "Cess%" : .ColHidden = True
                .Col = ENUM_Grid.VAT : .Text = "VAT%" : .ColHidden = True
            End If

            .Col = ENUM_Grid.ShopCode : .Text = "SHOP CODE"
            .Col = ENUM_Grid.UOM : .Text = "UOM"
            .Col = ENUM_Grid.ToolCost : .Text = "TOOL COST"
            .Col = ENUM_Grid.Packing_Per : .Text = "PACKING"
            .Col = ENUM_Grid.discount_type : .Text = "DISCOUNT TYPE"
            .Col = ENUM_Grid.discount_value : .Text = "DISCOUNT PER"
            .Col = ENUM_Grid.CGST_TAX : .Text = "CGST TAX"
            .Col = ENUM_Grid.SGST_TAX : .Text = "SGST TAX"
            .Col = ENUM_Grid.IGST_TAX : .Text = "IGST TAX"
            .Col = ENUM_Grid.COMP_ECC : .Text = "COMP Cess"
            .Col = ENUM_Grid.HSN_CODE : .Text = "HSN/SAC"
            .Col = ENUM_Grid.Remarks : .Text = "Remarks"
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
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Col = ENUM_Grid.ToolCost
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Col = ENUM_Grid.Packing_Per
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Col = ENUM_Grid.EDPer
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = ENUM_Grid.AddEx
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = ENUM_Grid.SHECess
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = ENUM_Grid.Cess
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = ENUM_Grid.VAT
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = ENUM_Grid.ShopCode
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.UOM
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            '10624527
            .Col = ENUM_Grid.discount_type
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = ENUM_Grid.discount_value
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            '10624527
            'GST CHANGES
            .Col = ENUM_Grid.CGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.SGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.IGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.COMP_ECC
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.HSN_CODE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

            'GST CHANGES
            .Col = ENUM_Grid.Remarks
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit

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
            .set_ColWidth(ENUM_Grid.Packing_Per, 5)
            .set_ColWidth(ENUM_Grid.ExWorks, 5)
            .set_ColWidth(ENUM_Grid.EDPer, 5)
            .set_ColWidth(ENUM_Grid.AddEx, 5)
            .set_ColWidth(ENUM_Grid.SHECess, 5)
            .set_ColWidth(ENUM_Grid.Cess, 5)
            .set_ColWidth(ENUM_Grid.VAT, 5)
            .set_ColWidth(ENUM_Grid.ShopCode, 8)
            .set_ColWidth(ENUM_Grid.UOM, 5)
            '10624527
            .set_ColWidth(ENUM_Grid.discount_type, 7)
            .set_ColWidth(ENUM_Grid.discount_value, 7)
            '10624527
            'GST CHANGES
            .set_ColWidth(ENUM_Grid.CGST_TAX, 10)
            .set_ColWidth(ENUM_Grid.SGST_TAX, 10)
            .set_ColWidth(ENUM_Grid.IGST_TAX, 10)
            .set_ColWidth(ENUM_Grid.COMP_ECC, 10)
            .set_ColWidth(ENUM_Grid.HSN_CODE, 20)
            'GST CHANGES
            .Col = ENUM_Grid.Remarks
            .set_ColWidth(ENUM_Grid.Remarks, 20)

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
    Public Sub CmdRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRefresh.Click
        Dim YesNo As String
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        YesNo = CStr(MsgBox("Do you want to Refresh?", MsgBoxStyle.YesNo, ResolveResString(100)))
        If YesNo = CStr(MsgBoxResult.Yes) Then
            Call RefreshForm()
        End If
        If YesNo = CStr(MsgBoxResult.No) Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'Changes against 10737738 
    Private Sub ChkVBSchUpdFlag()
        Dim strSql As String = String.Empty

        Try

            strSql = " select top 1 1 from sales_parameter where Unit_Code='" & gstrUNITID & "' and SCHEDULE_UPLOAD_CONFIG = 1  "
            SchUpdFlag = IsRecordExists(strSql)

        Catch ex As Exception
            Throw ex
        End Try

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
        Dim varDespatchQty As Object
        Dim varPrevQty As Object
        On Error GoTo ErrHandler
        If Len(Trim(txtDocNo.Text)) = 0 Then
            Exit Sub
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        txtQty.Text = CStr(0)
        txtSchQty.Text = CStr(0)
        If CmdUpdSchedule.Enabled = True Then
            rsWkngDays.GetResult("select distinct work_flg,dt from  calendar_mfg_mst where month (dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "'  and YEAR (dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' and dt > =  CONVERT(DATETIME,CONVERT(VARCHAR(10),getdate(),103),103) and UNIT_CODE='" & gstrUNITID & "' order by dt")
        Else
            rsWkngDays.GetResult("select distinct work_flg,dt from  calendar_mfg_mst where month (dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "'  and YEAR (dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' and dt > =   (select min(so_date) from so_upld_dtl where  doc_no = '" & txtDocNo.Text & "' and UNIT_CODE='" & gstrUNITID & "') and UNIT_CODE='" & gstrUNITID & "' order by dt")
        End If
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
        If SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) Then
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
        'INSERTION OF DATA IN DAILY MKT SCHEDULE GRID
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
        If Click_Renamed = True Then
            varItemCode = Nothing
            SSPurOrd.GetText(ENUM_Grid.InternalPartNo, SSPurOrd.ActiveRow, varItemCode)
        Else
            varItemCode = Nothing
            SSPurOrd.GetText(ENUM_Grid.InternalPartNo, NewRow, varItemCode)
        End If
        mP_Connection.Execute("set dateformat 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        rsSch.GetResult("select schedule_quantity,trans_date from dailymktschedule_temp  where item_code = '" & varItemCode & "'  and month(trans_date) = month('" & getDateForDB(dtpSOdate.Value) & "')  and year(trans_date) = year('" & getDateForDB(dtpSOdate.Value) & "') AND DOC_NO = '" & txtDocNo.Text & "' AND FILETYPE = 'SOUPLD'  and UNIT_CODE='" & gstrUNITID & "' order by trans_date")
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
    Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        On Error GoTo ErrHandler
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        Dim StrSql As String
        rssalesTerms = New ClsResultSetDB
        strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PY' and UNIT_CODE='" & gstrUNITID & "'"
        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rssalesTerms.GetNoRows > 0 Then
            strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PR' and UNIT_CODE='" & gstrUNITID & "'"
            rssalesTerms.ResultSetClose()
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows > 0 Then
                strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PK' and UNIT_CODE='" & gstrUNITID & "'"
                rssalesTerms.ResultSetClose()
                rssalesTerms = New ClsResultSetDB
                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalesTerms.GetNoRows > 0 Then
                    strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='FR' and UNIT_CODE='" & gstrUNITID & "'"
                    rssalesTerms.ResultSetClose()
                    rssalesTerms = New ClsResultSetDB
                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalesTerms.GetNoRows > 0 Then
                        strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='TR' and UNIT_CODE='" & gstrUNITID & "'"
                        rssalesTerms.ResultSetClose()
                        rssalesTerms = New ClsResultSetDB
                        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rssalesTerms.GetNoRows > 0 Then
                            strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='OC' and UNIT_CODE='" & gstrUNITID & "'"
                            rssalesTerms.ResultSetClose()
                            rssalesTerms = New ClsResultSetDB
                            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rssalesTerms.GetNoRows > 0 Then
                                strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='MO' and UNIT_CODE='" & gstrUNITID & "'"
                                rssalesTerms.ResultSetClose()
                                rssalesTerms = New ClsResultSetDB
                                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                If rssalesTerms.GetNoRows > 0 Then
                                    strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='DL' and UNIT_CODE='" & gstrUNITID & "'"
                                    rssalesTerms.ResultSetClose()
                                    rssalesTerms = New ClsResultSetDB
                                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rssalesTerms.GetNoRows > 0 Then
                                        'Select Case CmdButtons.mode
                                        frmMKTTRN0010.formload("MODE_ADD")
                                        Call frmMKTTRN0010.Show()
                                    Else
                                        Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                        cmdchangetype.Focus()
                                        Exit Sub
                                    End If
                                Else
                                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    cmdchangetype.Focus()
                                    Exit Sub
                                End If
                            Else
                                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                cmdchangetype.Focus()
                                Exit Sub
                            End If
                        Else
                            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            cmdchangetype.Focus()
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        cmdchangetype.Focus()
                        Exit Sub
                    End If
                Else
                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    cmdchangetype.Focus()
                    Exit Sub
                End If
            Else
                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                cmdchangetype.Focus()
                Exit Sub
            End If
        Else
            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            cmdchangetype.Focus()
            Exit Sub
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
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
        ''21478
        'Changes against 10737738 
        If SchUpdFlag = True Then
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select customer_Code,cust_name from customer_mst where UNIT_CODE='" & gstrUNITID & "' and SCH_UPLOAD_CODE ='SOUPLOAD' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", "Help", 2)
        Else
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select customer_Code,cust_name from customer_mst where UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", "Help", 2)
        End If
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
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "Excise Duty Help", 1)
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
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, " Select item_code ,item_desc ,cust_drgno,drg_desc   from custitem_mst where account_code = '" & txtCustomerCode.Text & "'  AND ACTIVE = 1 and UNIT_CODE='" & gstrUNITID & "'") ', "Help", 2)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
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
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "Additional Excise Duty Help", 1)
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
        Dim rsDocNo As New ClsResultSetDB
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        Call RefreshForm()
        StrSql = "SELECT cast(DOC_No as varchar(18)) as DOC_No ,CUSTOMER_CODE,Cust_Name FROM SO_UPLD_HDR SH  INNER JOIN CUSTOMER_MST CM ON CM.CUSTOMER_CODE = SH.CUST_CODE"
        StrSql = StrSql & " and CM.UNIT_CODE = SH.UNIT_CODE and CM.UNIT_CODE='" & gstrUNITID & "' and ((isnull(CM.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= CM.deactive_date))"
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "HELP", 1)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtDocNo.Text = strHelp(0)
                    docno = CInt(txtDocNo.Text)
                    mbln_View = True
                    Ctrl_EnabledDisable()
                    cmdSODetails.Enabled = True
                    cmdSODetails.Text = "Update SO"
                End If
            Else
                MsgBox(" No record available.", MsgBoxStyle.Information, ResolveResString(100))
                mbln_View = False
                cmdSODetails.Enabled = True
                cmdSODetails.Text = "Sales Order Details"
            End If
        End If
        If Len(txtDocNo.Text) > 0 Then
            docno = CInt(txtDocNo.Text)
            StrSql = "SELECT distinct H.DOC_NO,M.CUSTOMER_CODE,M.Cust_Name,H.REMARKS,D.VALID_FROM,  D.VALID_TO FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H  ON D.DOC_NO = H.DOC_NO AND D.UNIT_CODE = H.UNIT_CODE AND H.DOC_NO NOT IN (SELECT DISTINCT ISNULL(DOC_NO,0)  FROM dailymkTschedule WHERE FILETYPE = 'SOUPLD' AND UNIT_CODE='" & gstrUNITID & "') AND H.DOC_NO = '" & txtDocNo.Text & "'  INNER JOIN CUSTOMER_MST M ON  M.CUSTOMER_CODE = H.CUST_CODE "
            StrSql = StrSql & " AND M.UNIT_CODE = H.UNIT_CODE where D.UNIT_CODE='" & gstrUNITID & "'"
            rsDocNo.GetResult(StrSql)
            If rsDocNo.GetNoRows > 0 Then
                Call txtDocNo_Validating(txtDocNo, New System.ComponentModel.CancelEventArgs(True))
                cmdSODetails.Enabled = True
                cmdSODetails.Text = "Update SO"
                CmdUpdSchedule.Enabled = True
            Else
                Call txtDocNo_Validating(txtDocNo, New System.ComponentModel.CancelEventArgs(True))
                cmdSODetails.Enabled = False
                CmdUpdSchedule.Enabled = False
            End If
            rsDocNo.ResultSetClose()
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
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "Ecess Help", 1)
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
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "SHECess Help", 1)
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
        sql = "select distinct C.wh_code ,W.WH_Description  from custwarehouse_mst C, warehouse_mst W  where c.customer_code = '" & txtCustomerCode.Text & "'  and c.wh_code = w.wh_code"
        sql = sql & " and C.UNIT_CODE = W.UNIT_CODE and C.UNIT_CODE = '" & gstrUNITID & "' "
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
    Function Save_DailyMktSchedule() As Boolean
        On Error GoTo ErrHandler
        Dim varQty As Object
        Dim SCHQTY As Short
        Dim varPARTNO As Object
        Dim Col, Row As Short
        Dim rsWkngDays As New ClsResultSetDB
        Dim rsItemCode As New ClsResultSetDB
        Dim fraction As Short
        Dim intPos As Short
        Dim rsDespQty As ADODB.Recordset
        Dim rsPrevQty As ADODB.Recordset
        Dim PrevQty As Object = Nothing
        Dim DESPQty As Object = Nothing
        Dim varItemCode As Object = Nothing
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Dim rsMinQty As New ClsResultSetDB
        With SSPurOrd
            For Row = 1 To .MaxRows
                varQty = Nothing
                varPARTNO = Nothing
                .GetText(ENUM_Grid.Qty, Row, varQty)
                .GetText(ENUM_Grid.CustPartNo, Row, varPARTNO)
                rsDespQty = New ADODB.Recordset
                rsDespQty.Open("select isnull(sum(ISNULL(despatch_qty,0)),0) AS DESPATCH_QTY from dailymktschedule where cust_drgno = '" & varPARTNO & "' and account_code = '" & txtCustomerCode.Text & "' and trans_date <= CONVERT(DATETIME,CONVERT(VARCHAR(10),getdate(),103),103) and trans_date >= '" & getDateForDB(dtpDateFrom.Value) & "'  and schedule_flag = 1 and status = 1 and filetype = 'SOUPLD' AND UNIT_CODE = '" & gstrUNITID & "'  AND DOC_NO = (SELECT MAX(DOC_NO) FROM dailymktschedule WHERE ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND filetype = 'SOUPLD' AND DOC_NO < " & txtDocNo.Text & " AND UNIT_CODE='" & gstrUNITID & "')", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
                If rsDespQty.EOF = False And rsDespQty.BOF = False Then
                    DESPQty = rsDespQty.Fields("DESPATCH_QTY").Value
                    varQty = varQty - DESPQty
                Else
                    DESPQty = 0
                End If
                rsPrevQty = New ADODB.Recordset
                varItemCode = Nothing
                .GetText(ENUM_Grid.InternalPartNo, Row, varItemCode)
                rsPrevQty.Open("select isnull(sum(isnull(d.order_qty,0)),0) as PrevQty from cust_ord_hdr h, cust_ord_dtl d " & _
                    " where d.cust_ref = h.cust_ref and " & _
                    " h.amendment_no= d.amendment_no and" & _
                    " d.UNIT_CODE = h.UNIT_CODE and " & _
                    " h.account_code= d.account_code and" & _
                    " d.item_code = '" & varItemCode & "' and " & _
                    " h.account_code= '" & txtCustomerCode.Text & "'" & _
                    " and month(h.order_date) = " & Month(getDateForDB(dtpSOdate.Value)) & " " & _
                    " and year(h.order_date) = " & Year(getDateForDB(dtpSOdate.Value)) & " and d.UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
                If rsPrevQty.EOF = False And rsPrevQty.BOF = False Then
                    PrevQty = rsPrevQty.Fields("PrevQty").Value
                    varQty = varQty + PrevQty
                Else
                    PrevQty = 0
                End If
                mP_Connection.Execute("UPDATE SO_UPLD_DTL SET DESPATCHQTY = " & rsDespQty.Fields("DESPATCH_QTY").Value & ", PREVQTY = " & PrevQty & " WHERE DOC_NO = " & txtDocNo.Text & " AND PARTNO = '" & varPARTNO & "' and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                rsItemCode = New ClsResultSetDB
                rsItemCode.GetResult("SELECT DISTINCT ITEM_CODE  FROM CUSTITEM_MST  WHERE CUST_DRGNO = '" & varPARTNO & "' AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND ACTIVE = 1 and UNIT_CODE='" & gstrUNITID & "'")
                rsMinQty = New ClsResultSetDB
                rsMinQty.GetResult("select MinQtyOfSO from sales_parameter where UNIT_CODE='" & gstrUNITID & "'")
                If Val(Replace(varQty, ",", "")) <= rsMinQty.GetValue("MinQtyOfSO") Then
                    rsWkngDays.GetResult("select top 1 dt , work_flg from calendar_mfg_mst  where month(dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "' and  YEAR(dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' AND WORK_FLG = 0 AND DT > = CONVERT(DATETIME,CONVERT(VARCHAR(10),getdate(),103),103) and UNIT_CODE='" & gstrUNITID & "' order by dt")

                    mP_Connection.Execute("INSERT INTO DailyMktSchedule_Temp (Account_Code,  Trans_date,Item_code,Cust_Drgno,Schedule_Flag,  Schedule_Quantity,Despatch_Qty,Status,Ent_dt,  Ent_UserId,Upd_dt,Upd_UserId,Consignee_code,doc_no,filetype,UNIT_CODE)  VALUES('" & txtCustomerCode.Text & "','" & rsWkngDays.GetValue("dt") & "' ,  '" & rsItemCode.GetValue("ITEM_CODE") & "','" & varPARTNO & "',  1," & varQty & " ,0,0,GETDATE(),'" & mP_User & "',  GETDATE(),'" & mP_User & "','" & txtCustomerCode.Text & "','" & Trim(txtDocNo.Text) & "','SOUPLD','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Else
                    rsWkngDays.GetResult("select count(dt) as COUNTDT, month(dt) as month from calendar_mfg_mst  Where YEAR(dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' AND  work_flg = 0 and month(dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "' and dt > = CONVERT(DATETIME,CONVERT(VARCHAR(10),getdate(),103),103) and UNIT_CODE='" & gstrUNITID & "' group by month(dt)")
                    If rsWkngDays.GetNoRows > 0 Then
                        If rsWkngDays.GetValue("COUNTdt") > 0 Then
                            varQty = Val(Replace(varQty, ",", "")) / Val(rsWkngDays.GetValue("countdt"))
                            fraction = (varQty * 100) Mod 100
                            fraction = fraction * Val(rsWkngDays.GetValue("COUNTdt")) / 100
                            intPos = InStr(1, varQty, ".")
                            If intPos > 0 Then
                                varQty = VB.Left(varQty, intPos - 1)
                            End If
                            rsWkngDays.ResultSetClose()
                            rsWkngDays = New ClsResultSetDB
                            rsWkngDays.GetResult("select dt , work_flg from calendar_mfg_mst where month(dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "' and  YEAR(dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' AND WORK_FLG = 0 and dt>= CONVERT(DATETIME,CONVERT(VARCHAR(10),getdate(),103),103) and UNIT_CODE='" & gstrUNITID & "' order by dt")
                            If rsWkngDays.GetNoRows > 0 Then
                                rsWkngDays.MoveFirst()
                                While Not rsWkngDays.EOFRecord
                                    mP_Connection.Execute("INSERT INTO DailyMktSchedule_Temp (Account_Code,  Trans_date,Item_code,Cust_Drgno,Schedule_Flag,  Schedule_Quantity,Despatch_Qty,Status,Ent_dt,  Ent_UserId,Upd_dt,Upd_UserId,Consignee_code,doc_no,FILETYPE,UNIT_CODE)  VALUES('" & txtCustomerCode.Text & "','" & VB6.Format(rsWkngDays.GetValue("DT"), "dd MMM yyyy") & "',  '" & rsItemCode.GetValue("ITEM_CODE") & "',  '" & varPARTNO & "',  1," & varQty & " + " & fraction & ",0,1,GETDATE(),'" & mP_User & "',  GETDATE(),'" & mP_User & "','" & txtCustomerCode.Text & "','" & Trim(txtDocNo.Text) & "','SOUPLD','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    fraction = 0
                                    rsWkngDays.MoveNext()
                                End While
                            End If
                            rsItemCode.ResultSetClose()
                        End If
                    End If
                End If
                rsMinQty.ResultSetClose()
                rsMinQty = Nothing
            Next
            Save_DailyMktSchedule = True
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        Save_DailyMktSchedule = False
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Function
    Private Sub cmdSODetails_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSODetails.Click
        Dim MsgYesNo As String
        Dim Msg As String = ""
        Dim Row, Col As Short
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        If cmdSODetails.Text = "Update SO" Then
            update_SO()
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        If CheckUploaded_Data() = False Then
            txtQty.Text = ""
            txtSchQty.Text = ""
            txtDiff.Text = ""
            lblCustPartDesc.Text = ""
            lblItemDesc.Text = ""
            txtSearch.Text = ""
            For Row = 1 To 4
                For Col = 1 To 10
                    SpdSOSch.SetText(Col, Row, "")
                Next
            Next
            If txtCustomerCode.Text = "" Then
                Msg = Msg & vbCrLf & " Customer Code."
            End If
            'If txtPlantCode.Text = "" Then
            '    Msg = Msg & vbCrLf & " Plant Code."
            'End If
            If txtFileName.Text = "" Then
                Msg = Msg & vbCrLf & " File Name."
            End If
            If Len(LTrim(RTrim(Msg))) <> 0 Then
                MsgBox("Please Enter the following details: " & Msg, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Exit Sub
            End If
            Call SO_Upload()
        Else
            MsgYesNo = CStr(MsgBox("SO Details and Daily marketing Schedule Already Exists. Do You Want To Continue?", MsgBoxStyle.YesNo, ResolveResString(100)))
            If MsgYesNo = CStr(MsgBoxResult.Yes) Then
                Call SO_Upload()
            End If
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
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
        varMonthyear = Month(getDateForDB(dtpSOdate.Value)) & "/" & Year(getDateForDB(dtpSOdate.Value))
        Dim rsGetRows As New ClsResultSetDB
        If Val(txtDiff.Text) = 0 Then
            With SSPurOrd
                .Col = ENUM_Grid.InternalPartNo
                varItemCode = Nothing
                .GetText(ENUM_Grid.InternalPartNo, .ActiveRow, varItemCode)
                varPARTNO = Nothing
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
                                rsGetRows = New ClsResultSetDB
                                rsGetRows.GetResult("SELECT * FROM DAILYMKTSCHEDULE_TEMP  WHERE Account_Code = '" & txtCustomerCode.Text & "' and  Trans_date = '" & varDate & "' and  Item_code = '" & varItemCode & "'   AND FILETYPE = 'SOUPLD' and DOC_NO = '" & txtDocNo.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
                                If rsGetRows.GetNoRows > 0 Then
                                    mP_Connection.Execute("UPDATE DailyMktSchedule_Temp SET Schedule_Quantity = " & VarSchQty & "  WHERE Account_Code = '" & Trim(txtCustomerCode.Text) & "'  and Trans_date = '" & varDate & "' and  Item_code = '" & varItemCode & "' AND FILETYPE = 'SOUPLD'   AND DOC_NO = '" & txtDocNo.Text & "' and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                Else
                                    mP_Connection.Execute("INSERT INTO DailyMktSchedule_Temp (Account_Code,  Trans_date,Item_code,Cust_Drgno,Schedule_Flag,  Schedule_Quantity,Despatch_Qty,Status,Ent_dt,  Ent_UserId,Upd_dt,Upd_UserId,Consignee_code,doc_no,filetype,UNIT_CODE)  VALUES('" & txtCustomerCode.Text & "','" & varDate & "' ,  '" & varItemCode & "','" & varPARTNO & "',  1," & VarSchQty & " ,0,0,GETDATE(),'" & mP_User & "',  GETDATE(),'" & mP_User & "','" & txtCustomerCode.Text & "','" & Trim(txtDocNo.Text) & "','SOUPLD','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                                rsGetRows.ResultSetClose()
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
    Private Sub Upload_SO()
        On Error GoTo ErrHandler
        Dim varCustomerCode As String
        Dim Row As Short
        Dim varSONO As Object
        Dim varPARTNO As Object
        Dim varItemCode As Object
        Dim varPARTNAME As Object
        Dim varQty As Object
        Dim varEXWORKS As Object
        Dim varED As Object
        Dim varAddED As Object
        Dim varSHECess As Object
        Dim varCess As Object
        Dim varVAT As Object
        Dim varCustSupply As Object
        Dim varToolCost As Object
        Dim varpackingper As Object
        Dim varRemarks As Object
        Dim blnflag As Boolean
        Dim Msg As String = ""
        Dim StrSql As String = ""
        Dim Rsdoc As New ClsResultSetDB
        Dim rsDuplicate As New ClsResultSetDB
        Dim rsDespQty As ADODB.Recordset
        Dim rsPrevQty As ADODB.Recordset
        Dim rsPrevDocNo As ADODB.Recordset
        Dim rsIsAuth As ADODB.Recordset
        Dim PrevQty As Object = Nothing
        Dim DESPQty As Object = Nothing
        Dim decexworks As Decimal = 0.0
        Dim varSHOPCODE As Object
        Dim varUOM As Object
        Dim VARINTERNALSALESORDER As String = String.Empty
        Dim strTempSeries As String
        Dim strCheckDOcNo As String
        Dim intMaxLoop As Short
        Dim strZeroSuffix As String
        Dim STRDOCCNO As String
        Dim intLoopCounter As Short
        Dim strCT2Condition As String = ""
        Dim strMsg As String = ""
        Dim StrCT2Reqd As String = ""
        Dim strCustPartNo As String = String.Empty
        Dim blnIsEopRequired As Boolean = False
        Dim intRow As Short
        'GST CHANGES
        Dim VARHSNSACCODE As Object
        Dim VARCGSTTAX As Object
        Dim VARSGSTTAX As Object
        Dim VARIGSTTAX As Object
        Dim VARCOMPECC As Object
        Dim VARCGSTTAX_TYPE As Object
        Dim VARSGSTTAX_TYPE As Object
        Dim VARIGSTTAX_TYPE As Object
        Dim VARCOMPECC_TYPE As Object
        Dim STRHSNCODE As String
        Dim MsgYesNo As String
        Dim strReturnCustRef As String
        Dim blnopenso As Boolean

        Dim objSQLConn As SqlConnection
        Dim objReader As SqlDataReader
        Dim objCommand As SqlCommand
        Dim ISHSNORSAC As String
        Dim HSNSACCODE As String
        Dim CGST_TXRT_HEAD As String
        Dim SGST_TXRT_HEAD As String
        Dim UGST_TXRT_HEAD As String
        Dim IGST_TXRT_HEAD As String
        Dim COMPENSATION_CESS As String

        'GST CHANGES

        '10736222-changes done for CT2 ARE 3 functionality
        If ChkCT2reqd.Checked = True Then
            If SSPurOrd.MaxRows > 0 Then
                Row = 1
                While Row <= SSPurOrd.MaxRows
                    With SSPurOrd
                        varItemCode = Nothing
                        varPARTNO = Nothing
                        VARHSNSACCODE = Nothing

                        .GetText(ENUM_Grid.InternalPartNo, Row, varItemCode)
                        .GetText(ENUM_Grid.CustPartNo, Row, varPARTNO)

                        VARHSNSACCODE = Find_Value("SELECT hsn_sac_code FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & varItemCode & "'")
                        StrSql = "SELECT * FROM CT2_CUST_ITEM_LINKAGE WHERE Unit_code='" & gstrUNITID & "' AND ITEM_CODE='" & varItemCode & "' AND CUST_DRGNO='" & varPARTNO & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "' AND ISAUTHORIZED=1 AND ACTIVE=1 "
                        If IsRecordExists(StrSql) = False Then
                            strMsg = strMsg + varItemCode + " - " + varPARTNO + ", "
                        End If

                        
                    End With
                    Row = Row + 1
                End While

                If strMsg.Trim.Length > 0 Then
                    MessageBox.Show("CT2 Item Linkage not found for below parts : " & vbCr & strMsg, ResolveResString(100))
                    Exit Sub
                End If
            End If
        End If

        If gblnGSTUnit = True Then
            If SSPurOrd.MaxRows > 0 Then
                Row = 1
                While Row <= SSPurOrd.MaxRows
                    With SSPurOrd
                        varItemCode = Nothing
                        VARHSNSACCODE = Nothing

                        .GetText(ENUM_Grid.InternalPartNo, Row, varItemCode)

                        VARHSNSACCODE = Find_Value("SELECT hsn_sac_code FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & varItemCode & "'")
                        StrSql = "SELECT * FROM ITEM_MST  WHERE Unit_code='" & gstrUNITID & "' AND ITEM_CODE='" & varItemCode & "' and isnull(HSN_SAC_CODE,'')=''  "
                        If IsRecordExists(StrSql) = True Then
                            strMsg = strMsg + varItemCode + ", "
                        End If
                    End With
                    Row = Row + 1
                End While

                If strMsg.Trim.Length > 0 Then
                    MessageBox.Show("HSN/SAC CODE NOT FOUND FOR BELOW ITEM CODES : " & vbCr & strMsg, ResolveResString(100))
                    Exit Sub
                End If
            End If



        End If


        '10808160--Starts
        StrSql = "Select dbo.UDF_ISEOPREQUIRED('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "')"
        blnIsEopRequired = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(StrSql))
        If blnIsEopRequired = True Then
            With SSPurOrd
                For intRow = 1 To SSPurOrd.MaxRows
                    varItemCode = Nothing
                    varPARTNO = Nothing

                    .GetText(ENUM_Grid.InternalPartNo, intRow, varItemCode)
                    .GetText(ENUM_Grid.CustPartNo, intRow, varPARTNO)

                    StrSql = "SELECT TOP 1 CUST_DRGNO FROM BUDGETITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & Trim(txtCustomerCode.Text) & "' " & _
                                    " AND ITEM_CODE = '" & Trim(varItemCode) & "' AND CUST_DRGNO = '" & Trim(varPARTNO) & "' AND ENDDATE < '" & Convert.ToDateTime(dtpDateTo.Value).ToString("dd MMM yyyy") & "' "
                    If IsRecordExists(StrSql) = True Then
                        strCustPartNo = strCustPartNo + varPARTNO + " , "
                    End If
                Next

                If strCustPartNo <> "" Then
                    If MsgBox("End Date of Following Cust. Part: " + " '" & Trim(strCustPartNo) & "' " + "  " & vbCrLf & "In this SO are falling before the Validity Date of the SO." & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End With
        End If
        '10808160--Ends

        If blnlinelevelcustomer = True Then
            VARINTERNALSALESORDER = txtCustomerCode.Text & "-" & VB6.Format(dtpSOdate.Value, "MMYY") & "-"
            STRDOCCNO = Find_Value("SELECT SERIES_FOR_LINELEVELSO FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'")

            strTempSeries = CStr(CInt(STRDOCCNO) + 1)

            If Len(Trim(strTempSeries)) < 6 Then
                intMaxLoop = 6 - Len(Trim(strTempSeries))
                strZeroSuffix = ""
                For intLoopCounter = 1 To intMaxLoop
                    strZeroSuffix = Trim(strZeroSuffix) & "0"
                Next
            End If
            VARINTERNALSALESORDER = VARINTERNALSALESORDER + strZeroSuffix + strTempSeries

        End If

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Rsdoc.GetResult("SELECT CURRENT_NO + 1 AS DOC_NO FROM DOCUMENTTYPE_MST  WHERE DOC_TYPE = 304 and UNIT_CODE='" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        If txtCustomerCode.Text = "" Then
            Msg = Msg & vbCrLf & " Customer Code."
        End If
        'If txtPlantCode.Text = "" Then
        '    Msg = Msg & vbCrLf & " Plant Code."
        'End If
        If txtFileName.Text = "" Then
            Msg = Msg & vbCrLf & " File Name."
        End If

        
        If Len(LTrim(RTrim(Msg))) <> 0 Then
            MsgBox("Please Enter the following details: " & Msg, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        If txtECESS.Text = "" Then
            txtECESS.Text = CStr(0)
        End If
        If txtSECESS.Text = "" Then
            txtSECESS.Text = CStr(0)
        End If

        If txtShipAddress.Text = "" Then
            MsgYesNo = CStr(MsgBox("Do You Want To upload file without Ship address Code?", MsgBoxStyle.YesNo, ResolveResString(100)))
            If MsgYesNo = CStr(MsgBoxResult.No) Then
                mblnshippingcodemandatory = True
                Exit Sub
            End If
        End If


        If Rsdoc.GetNoRows > 0 Then
            mP_Connection.BeginTrans()
            Dim vbYesNo As String
            rsPrevDocNo = New ADODB.Recordset
            rsPrevDocNo.Open("select max(doc_no) as doc_no from so_upld_dtl where Month(so_date) = " & Month(getDateForDB(dtpSOdate.Value)) & " " & _
            " And Year(so_date) = " & Year(getDateForDB(dtpSOdate.Value)) & " and doc_no < " & Rsdoc.GetValue("doc_no") & " and UNIT_CODE='" & gstrUNITID & "' ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
            If Not IsDBNull(rsPrevDocNo.Fields("doc_no").Value) Then
                rsIsAuth = New ADODB.Recordset
                rsIsAuth.Open("select doc_no from dailymktschedule where doc_no = '" & rsPrevDocNo.Fields("doc_no").Value & "'  and filetype = 'SOUPLD' and UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
                If rsIsAuth.EOF = True And rsIsAuth.BOF = True Then
                    vbYesNo = MsgBox("Schedule For The YearMonth " & CStr(Year(getDateForDB(dtpSOdate.Value))) & CStr(Month(getDateForDB(dtpSOdate.Value))) & vbCrLf & "With Doc No " & rsPrevDocNo.Fields("doc_no").Value & vbCrLf & " is not Authorised." + vbCrLf + "Authorize That Schedule First, Otherwise You" + vbCrLf + "Can't Use That Schedule Further." & vbCrLf + "Do You Want To Continue?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, ResolveResString(100))
                    If vbYesNo = CStr(MsgBoxResult.No) Then
                        mP_Connection.RollbackTrans()
                        Exit Sub
                    End If
                    If vbYesNo = CStr(MsgBoxResult.Yes) Then
                        mP_Connection.Execute("UPDATE SO_UPLD_HDR SET CANAUTH = 0 WHERE DOC_NO = '" & rsPrevDocNo.Fields("doc_no").Value & "' and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        frmMKTTRN0058.Close()
                    End If
                End If
            End If
            
            '10736222-changes done for CT2 ARE 3 functionality
            If ChkCT2reqd.CheckState = System.Windows.Forms.CheckState.Checked Then
                If mblnSOUPLD_ForexCurrency = True Then
                    'mP_Connection.Execute("Insert Into SO_Upld_Hdr (cust_code,plant_c,file_location,so_date,remarks,Ent_Uid,Doc_no,UNIT_CODE,CT2_Reqd_InSO,SOUPLD_Forexcurrency,SHIPADDRESS_CODE,SHIPADDRESS_DESC )  Values('" & txtCustomerCode.Text & "','" & txtPlantCode.Text & "',   '" & Replace(txtFileName.Text, "'", "''") & "' ,'" & getDateForDB(dtpSOdate.Value) & "', '" & Replace(txtRemarks.Text, "'", "''") & "',  '" & mP_User & "'," & Rsdoc.GetValue("DOC_NO") & ",'" & gstrUNITID & "','1',1)", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute("Insert Into SO_Upld_Hdr (cust_code,plant_c,file_location,so_date,remarks,Ent_Uid,Doc_no,UNIT_CODE,CT2_Reqd_InSO,SOUPLD_Forexcurrency,SHIPADDRESS_CODE,SHIPADDRESS_DESC )  Values('" & txtCustomerCode.Text & "','" & txtPlantCode.Text & "',   '" & Replace(txtFileName.Text, "'", "''") & "' ,'" & getDateForDB(dtpSOdate.Value) & "', '" & Replace(txtRemarks.Text, "'", "''") & "',  '" & mP_User & "'," & Rsdoc.GetValue("DOC_NO") & ",'" & gstrUNITID & "','1',1,'" & txtShipAddress.Text & "','" & lblShipAddress_Details.Text & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Else
                    mP_Connection.Execute("Insert Into SO_Upld_Hdr (cust_code,plant_c,file_location,so_date,remarks,Ent_Uid,Doc_no,UNIT_CODE,CT2_Reqd_InSO,SOUPLD_Forexcurrency )  Values('" & txtCustomerCode.Text & "','" & txtPlantCode.Text & "',   '" & Replace(txtFileName.Text, "'", "''") & "' ,'" & getDateForDB(dtpSOdate.Value) & "', '" & Replace(txtRemarks.Text, "'", "''") & "',  '" & mP_User & "'," & Rsdoc.GetValue("DOC_NO") & ",'" & gstrUNITID & "','1',0,'" & txtShipAddress.Text & "','" & lblShipAddress_Details.Text & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If

            Else
                If mblnSOUPLD_ForexCurrency = True Then
                    mP_Connection.Execute("Insert Into SO_Upld_Hdr (cust_code,plant_c,file_location,so_date,remarks,Ent_Uid,Doc_no,UNIT_CODE,CT2_Reqd_InSO,SOUPLD_Forexcurrency ,SHIPADDRESS_CODE,SHIPADDRESS_DESC)  Values('" & txtCustomerCode.Text & "','" & txtPlantCode.Text & "',   '" & Replace(txtFileName.Text, "'", "''") & "' ,'" & getDateForDB(dtpSOdate.Value) & "', '" & Replace(txtRemarks.Text, "'", "''") & "',  '" & mP_User & "'," & Rsdoc.GetValue("DOC_NO") & ",'" & gstrUNITID & "','0',1,'" & txtShipAddress.Text & "','" & lblShipAddress_Details.Text & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Else
                    mP_Connection.Execute("Insert Into SO_Upld_Hdr (cust_code,plant_c,file_location,so_date,remarks,Ent_Uid,Doc_no,UNIT_CODE,CT2_Reqd_InSO,SOUPLD_Forexcurrency,SHIPADDRESS_CODE,SHIPADDRESS_DESC )  Values('" & txtCustomerCode.Text & "','" & txtPlantCode.Text & "',   '" & Replace(txtFileName.Text, "'", "''") & "' ,'" & getDateForDB(dtpSOdate.Value) & "', '" & Replace(txtRemarks.Text, "'", "''") & "',  '" & mP_User & "'," & Rsdoc.GetValue("DOC_NO") & ",'" & gstrUNITID & "','0',0, '" & txtShipAddress.Text & "','" & lblShipAddress_Details.Text & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If

            End If
        Else
            MsgBox("Doc Type Not Defined in DocumentType_Mst", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
        End If

        Dim vardISCOUNTTYPE As Object
        Dim vardISCOUNTPER As Object
        If SSPurOrd.MaxRows > 0 Then
            Row = 1
            If Rsdoc.GetNoRows > 0 Then
                While Row <= SSPurOrd.MaxRows
                    With SSPurOrd
                        varSONO = Nothing
                        varItemCode = Nothing
                        varPARTNO = Nothing
                        varPARTNAME = Nothing
                        varQty = Nothing
                        varEXWORKS = Nothing
                        varED = Nothing
                        varAddED = Nothing
                        varSHECess = Nothing
                        varCess = Nothing
                        varVAT = Nothing
                        varCustSupply = Nothing
                        varToolCost = Nothing
                        varpackingper = Nothing
                        varRemarks = Nothing
                        varSHOPCODE = Nothing
                        varUOM = Nothing
                        vardISCOUNTTYPE = Nothing
                        vardISCOUNTPER = Nothing
                        VARCGSTTAX = Nothing
                        VARSGSTTAX = Nothing
                        VARIGSTTAX = Nothing
                        VARCOMPECC = Nothing

                        .GetText(ENUM_Grid.salesorder, Row, varSONO)
                        .GetText(ENUM_Grid.InternalPartNo, Row, varItemCode)
                        .GetText(ENUM_Grid.CustPartNo, Row, varPARTNO)
                        .GetText(ENUM_Grid.CustPartDesc, Row, varPARTNAME)
                        .GetText(ENUM_Grid.Qty, Row, varQty)
                        .GetText(ENUM_Grid.ExWorks, Row, varEXWORKS)
                        .GetText(ENUM_Grid.EDPer, Row, varED)
                        .GetText(ENUM_Grid.AddEx, Row, varAddED)
                        .GetText(ENUM_Grid.SHECess, Row, varSHECess)
                        .GetText(ENUM_Grid.Cess, Row, varCess)
                        .GetText(ENUM_Grid.VAT, Row, varVAT)
                        .GetText(ENUM_Grid.custsupply, Row, varCustSupply)
                        .GetText(ENUM_Grid.ToolCost, Row, varToolCost)
                        .GetText(ENUM_Grid.Packing_Per, Row, varpackingper)
                        .GetText(ENUM_Grid.ShopCode, Row, varSHOPCODE)
                        .GetText(ENUM_Grid.UOM, Row, varUOM)
                        .GetText(ENUM_Grid.Remarks, Row, varRemarks)
                        .GetText(ENUM_Grid.CGST_TAX, Row, VARCGSTTAX)
                        .GetText(ENUM_Grid.SGST_TAX, Row, VARSGSTTAX)
                        .GetText(ENUM_Grid.IGST_TAX, Row, VARIGSTTAX)
                        .GetText(ENUM_Grid.COMP_ECC, Row, VARCOMPECC)
                        .GetText(ENUM_Grid.discount_type, Row, vardISCOUNTTYPE)
                        .GetText(ENUM_Grid.discount_value, Row, vardISCOUNTPER)

                        ''added to resolve tax issue on 26 Dec 2023 
                        VARHSNSACCODE = Find_Value("SELECT hsn_sac_code FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & varItemCode & "'")
                        ''end to resolve tax issue on 26 Dec 2023 

                        VARCGSTTAX_TYPE = Find_Value("SELECT TOP 1 TxRt_Rate_No  FROM Gen_TaxRate WHERE Tx_TaxeID='CGST' AND UNIT_CODE ='" & gstrUNITID & "' AND TxRt_Percentage ='" & VARCGSTTAX & "'")
                        VARSGSTTAX_TYPE = Find_Value("SELECT TOP 1 TxRt_Rate_No  FROM Gen_TaxRate WHERE Tx_TaxeID='SGST' AND UNIT_CODE ='" & gstrUNITID & "' AND TxRt_Percentage ='" & VARSGSTTAX & "'")
                        VARIGSTTAX_TYPE = Find_Value("SELECT TOP 1 TxRt_Rate_No  FROM Gen_TaxRate WHERE Tx_TaxeID='IGST' AND UNIT_CODE ='" & gstrUNITID & "' AND TxRt_Percentage ='" & VARIGSTTAX & "'")
                        VARCOMPECC_TYPE = Find_Value("SELECT TOP 1 TxRt_Rate_No  FROM Gen_TaxRate WHERE Tx_TaxeID='COMPECC' AND UNIT_CODE ='" & gstrUNITID & "' AND TxRt_Percentage ='" & VARCOMPECC & "'")

                        'If gblnGSTUnit = True Then '' Added by priti to check tax mapping INC1626086 
                        '    CGST_TXRT_HEAD = Nothing
                        '    SGST_TXRT_HEAD = Nothing
                        '    IGST_TXRT_HEAD = Nothing

                        '    StrSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "','" & varItemCode & "','','')"
                        '    Dim dtTax As DataTable = SqlConnectionclass.GetDataTable(StrSql)
                        '    If dtTax.Rows.mpuCount > 0 Then
                        '        ISHSNORSAC = Convert.ToString(dtTax.Rows(0)(0))
                        '        HSNSACCODE = Convert.ToString(dtTax.Rows(0)(1))
                        '        CGST_TXRT_HEAD = Convert.ToString(dtTax.Rows(0)(2))
                        '        SGST_TXRT_HEAD = Convert.ToString(dtTax.Rows(0)(3))
                        '        UGST_TXRT_HEAD = Convert.ToString(dtTax.Rows(0)(4))
                        '        IGST_TXRT_HEAD = Convert.ToString(dtTax.Rows(0)(5))
                        '        COMPENSATION_CESS = Convert.ToString(dtTax.Rows(0)(6))
                        '    End If
                        '    If Len(CGST_TXRT_HEAD) > 0 Or VARCGSTTAX > 0 Then
                        '        If (VARCGSTTAX_TYPE <> CGST_TXRT_HEAD) Or (VARSGSTTAX_TYPE <> SGST_TXRT_HEAD) Then
                        '            MsgBox("Check Tax Mapping at Item " & varItemCode, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        '            mP_Connection.RollbackTrans()
                        '            Exit Sub
                        '        End If
                        '    End If

                        '    If Len(IGST_TXRT_HEAD) > 0 Or VARIGSTTAX > 0 Then
                        '        If VARIGSTTAX_TYPE <> IGST_TXRT_HEAD Then
                        '            MsgBox("Check Tax Mapping at Item " & varItemCode, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        '            mP_Connection.RollbackTrans()
                        '            Exit Sub
                        '        End If
                        '    End If
                        'End If

                        If Val(varCustSupply).ToString = "0" Then
                            varCustSupply = 0
                        End If
                        If Val(varToolCost).ToString = "0" Then
                            varToolCost = 0
                        End If
                        If Val(varpackingper).ToString = "0" Then
                            varpackingper = 0
                        End If

                        If varSHECess = "" Then
                            varSHECess = 0
                        End If
                        If varQty = 0 Then
                            blnopenso = 1
                        Else
                            blnopenso = 0
                        End If
                        strReturnCustRef = CheckForMultipleOpenSO(txtCustomerCode.Text, txtCustomerCode.Text, varSONO, varPARTNO, varItemCode, blnopenso)
                        If Len(strReturnCustRef) > 0 Then
                            MsgBox("More than one Sales Order cannot be active for Customer Item Combination" & strReturnCustRef & "Part No :" & varPARTNO, vbInformation, ResolveResString(100))
                            Exit Sub
                        End If '10624527
                        If gblnGSTUnit = False AndAlso vardISCOUNTTYPE.ToString.Length > 0 Then
                            If Not (UCase(vardISCOUNTTYPE.ToString) = "P" Or UCase(vardISCOUNTTYPE.ToString) = "V") Then
                                MsgBox("DISCOUNT TYPE SHOULD BE EITHER P OR V ." & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                mP_Connection.RollbackTrans()
                                cmdSODetails.Enabled = False
                                CmdUpdSchedule.Enabled = False
                                Exit Sub
                            End If

                            If UCase(vardISCOUNTTYPE.ToString) = "P" Then
                                If vardISCOUNTPER > 100 Or vardISCOUNTPER < 0 Then
                                    MsgBox("PERCENTAGE SHOULD BE IN BETWEEN 0-100 " & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                    mP_Connection.RollbackTrans()
                                    cmdSODetails.Enabled = False
                                    CmdUpdSchedule.Enabled = False
                                    Exit Sub
                                End If
                            End If

                            If UCase(vardISCOUNTTYPE.ToString) = "V" Then
                                If vardISCOUNTPER > varEXWORKS Then
                                    MsgBox("DISCOUNT VALUE SHOULD NOT BE GREATER THAN RATE " & varEXWORKS & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                    mP_Connection.RollbackTrans()
                                    cmdSODetails.Enabled = False
                                    CmdUpdSchedule.Enabled = False
                                    Exit Sub
                                End If
                            End If

                        End If

                    End With
                    Dim vardISCOUNTTYPEDESC As String
                    txtDocNo.Text = Rsdoc.GetValue("doc_no ")
                    If UCase(vardISCOUNTTYPE) = "P" Then
                        vardISCOUNTTYPEDESC = "[P]ercentage"
                    End If
                    If UCase(vardISCOUNTTYPE) = "V" Then
                        vardISCOUNTTYPEDESC = "Value"
                    End If
                    '10624527
                    If blnSOUPLD_NotRemovehypen = True Then
                        varPARTNO = varPARTNO
                    Else
                        varPARTNO = Replace(varPARTNO, "-", "")
                    End If
                    rsPrevQty = New ADODB.Recordset
                    rsPrevQty.Open("select isnull(sum(isnull(d.order_qty,0)),0) as PrevQty from cust_ord_hdr h, cust_ord_dtl d " & _
                    " where d.cust_ref = h.cust_ref and" & _
                    " h.amendment_no= d.amendment_no and" & _
                    " h.account_code= d.account_code and" & _
                    " d.UNIT_CODE = h.UNIT_CODE and " & _
                    " d.item_code = '" & varItemCode & "' and " & _
                    " h.account_code= '" & txtCustomerCode.Text & "'" & _
                    " and month(h.order_date) = " & Month(getDateForDB(dtpSOdate.Value)) & " " & _
                    " and year(h.order_date) = " & Year(getDateForDB(dtpSOdate.Value)) & " and d.UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
                    If Not rsPrevQty.EOF And Not rsPrevQty.BOF Then
                        PrevQty = rsPrevQty.Fields("PrevQty").Value
                    Else
                        PrevQty = 0
                    End If

                    SSPurOrd.Col = ENUM_Grid.PrevQty
                    SSPurOrd.Row = Row
                    SSPurOrd.Text = CInt(PrevQty)
                    rsDespQty = New ADODB.Recordset
                    rsDespQty.Open("select isnull(sum(ISNULL(despatch_qty,0)),0) AS DESPATCH_QTY from dailymktschedule where cust_drgno = '" & varPARTNO & "' and account_code = '" & txtCustomerCode.Text & "' and trans_date <= CONVERT(DATETIME,CONVERT(VARCHAR(10),getdate(),103),103) and trans_date >= '" & getDateForDB(dtpDateFrom.Value) & "' and schedule_flag = 1 and status = 1 and filetype = 'SOUPLD' and UNIT_CODE='" & gstrUNITID & "' AND DOC_NO = (SELECT MAX(DOC_NO) FROM dailymktschedule WHERE ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND filetype = 'SOUPLD' AND DOC_NO < " & Rsdoc.GetValue("DOC_NO") & " and UNIT_CODE='" & gstrUNITID & "')", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
                    If Not rsDespQty.EOF And Not rsDespQty.BOF Then
                        DESPQty = rsDespQty.Fields("DESPATCH_QTY").Value
                    Else
                        DESPQty = 0
                    End If
                    SSPurOrd.Col = ENUM_Grid.DespatchQty
                    SSPurOrd.Row = Row

                    SSPurOrd.Text = CInt(DESPQty)
                    decexworks = 0.0
                    If varEXWORKS <> "" Then
                        decexworks = Decimal.Parse(varEXWORKS)
                    End If
                    StrSql = "INSERT INTO SO_UPLD_DTL (CUST_CODE,SalesOrder,Valid_From,Valid_to,So_Date,"
                    StrSql = StrSql & " item_code,PartNo,PartName,"
                    StrSql = StrSql & " Quantity,DespatchQty,PrevQty,EDPer,ExWorks,SHECess,Cess,VAT,CustSupply,ToolCost,packing_per,"
                    StrSql = StrSql & " Remarks, Doc_No, AddEx, Currency_Code,"
                    StrSql = StrSql & " Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,"
                    StrSql = StrSql & " Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,"
                    StrSql = StrSql & " Mode_Despatch,Delivery,"
                    StrSql = StrSql & " Consignee_code,Ent_Uid,Shop_code,uom,UNIT_CODE,INTERNAL_SALESORDER_NO,DISCOUNT_TYPE,DISCOUNT_VALUE"
                    If gblnGSTUnit = True Then
                        StrSql = StrSql & " ,HSNSACCODE,CGST_TAX,SGST_TAX,IGST_TAX,COMPCC_TAX)"
                    Else
                        StrSql = StrSql & ")"
                    End If
                    StrSql = StrSql & " VALUES( '" & txtCustomerCode.Text & "','" & Replace(varSONO, "'", "''") & "',"
                    StrSql = StrSql & " '" & getDateForDB(dtpDateFrom.Value) & "','" & getDateForDB(dtpDateTo.Value) & "','" & getDateForDB(dtpSOdate.Value) & "',"
                    StrSql = StrSql & "'" & Trim(varItemCode) & "','" & Replace(varPARTNO, "'", "''") & "',"
                    StrSql = StrSql & "'" & Replace(varPARTNAME, "'", "''") & "','" & Replace(varQty, ",", "") & "','" & DESPQty & "',"
                    StrSql = StrSql & "'" & PrevQty & "', '" & LTrim(RTrim(varED)) & "'," & decexworks & ",'" & LTrim(RTrim(varSHECess)) & "',"
                    StrSql = StrSql & "'" & LTrim(RTrim(varCess)) & "','" & LTrim(RTrim(varVAT)) & "', '" & Replace(varCustSupply, "'", "''") & "',"
                    StrSql = StrSql & "" & varToolCost & "," & varpackingper & ",'" & Replace(varRemarks, "'", "''") & "',"
                    StrSql = StrSql & "'" & Rsdoc.GetValue("DOC_NO") & "','" & LTrim(RTrim(txtAddExcise.Text)) & "','" & Trim(LblCurrencyCode.Text) & "',"
                    StrSql = StrSql & "'" & Trim(lblCredit_days.Text) & "','" & Trim(m_strSpecialNotes) & "','" & Trim(m_strPaymentTerms) & "',"
                    StrSql = StrSql & "'" & Trim(m_strPricesAre) & "','" & Trim(m_strPkgAndFwd) & "',"
                    StrSql = StrSql & "'" & Trim(m_strFreight) & "','" & Trim(m_strTransitInsurance) & "',"
                    StrSql = StrSql & "'" & Trim(m_strOctroi) & "','" & Trim(m_strModeOfDespatch) & "',"
                    '10624527
                    If blnlinelevelcustomer = True Then
                        'StrSql = StrSql & "'" & Trim(m_strDeliverySchedule) & "','" & Trim(txtCustomerCode.Text) & "','" & mP_User & "','" & varSHOPCODE & "','" & varUOM & "','" & gstrUNITID & "','" & VARINTERNALSALESORDER & "')"
                        StrSql = StrSql & "'" & Trim(m_strDeliverySchedule) & "','" & Trim(txtCustomerCode.Text) & "','" & mP_User & "','" & varSHOPCODE & "','" & varUOM & "','" & gstrUNITID & "','" & VARINTERNALSALESORDER & "','" & vardISCOUNTTYPE & "','" & vardISCOUNTPER & "'"
                    Else
                        StrSql = StrSql & "'" & Trim(m_strDeliverySchedule) & "','" & Trim(txtCustomerCode.Text) & "','" & mP_User & "','" & varSHOPCODE & "','" & varUOM & "','" & gstrUNITID & "','" & Replace(varSONO, "'", "''") & "','" & vardISCOUNTTYPEDESC & "','" & vardISCOUNTPER & "'"
                    End If
                    '10624527
                    If gblnGSTUnit = True Then
                        StrSql = StrSql & ",'" & VARHSNSACCODE & "','" & VARCGSTTAX_TYPE & "','" & VARSGSTTAX_TYPE & "','" & VARIGSTTAX_TYPE & "','" & VARCOMPECC_TYPE & "')"
                    Else
                        StrSql = StrSql & ")"
                    End If
                    mP_Connection.Execute(StrSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    
                    Row = Row + 1
                End While
            End If
        End If
        '10736222-changes done for CT2 ARE 3 functionality
        'Valid From and Valid To validation added in query by Prashant Rajpal on 03 Oct 2018
        rsDuplicate.GetResult("SELECT A.CUST_CODE,SalesOrder, PARTNO,count(PARTNO) as count, count(salesorder) as countSO FROM SO_Upld_dtl A, SO_Upld_Hdr B WHERE A.Cust_Code = B.Cust_Code and A.Doc_No =  B.Doc_No and A.Unit_Code =  B.Unit_Code and A.CUST_CODE = '" & txtCustomerCode.Text & "' and A.UNIT_CODE='" & gstrUNITID & "' and B.CanAuth=0 and convert(varchar(11),getdate(),106) between valid_from and valid_to group by A.CUST_CODE,SalesOrder,PARTNO")
        Dim DuplicateItem As String = ""
        If rsDuplicate.GetNoRows > 0 Then
            rsDuplicate.MoveFirst()
            While Not rsDuplicate.EOFRecord
                If rsDuplicate.GetValue("count") > 1 Or rsDuplicate.GetValue("countSO") > 1 Then
                    DuplicateItem = DuplicateItem + rsDuplicate.GetValue("PARTNO") + " : " + rsDuplicate.GetValue("SalesOrder") + vbCrLf
                    rsDuplicate.MoveNext()
                Else
                    rsDuplicate.MoveNext()
                End If
            End While
        End If
        rsDuplicate.ResultSetClose()
        If Len(Trim(DuplicateItem)) > 0 And mblnDUPLICATE_SOALLOWED = False Then
            MsgBox("There are Following Duplicate Entries : " & vbCrLf & "Item Code : Sales Order NO." & vbCrLf & DuplicateItem & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            mP_Connection.RollbackTrans()
            cmdSODetails.Enabled = False
            CmdUpdSchedule.Enabled = False
            Exit Sub
        End If
        txtDocNo.Text = Rsdoc.GetValue("doc_no ")
        If Save_DailyMktSchedule() = True Then
            mP_Connection.Execute("UPDATE DOCUMENTTYPE_MST SET CURRENT_NO = CURRENT_NO + 1 WHERE DOC_TYPE = 304 and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If blnlinelevelcustomer = True Then
                mP_Connection.Execute("UPDATE SALES_PARAMETER SET SERIES_FOR_LINELEVELSO  = SERIES_FOR_LINELEVELSO  + 1 WHERE UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If mblnUPDATE_DMS_SO_SCH_UPLD =True 
                    MsgBox("SO NO : " & VARINTERNALSALESORDER & " Details and Daily Marketing Schedule uploaded. ", MsgBoxStyle.OkOnly, ResolveResString(100))
                Else
                    MsgBox("SO NO : " & VARINTERNALSALESORDER & " Details uploaded. ", MsgBoxStyle.OkOnly, ResolveResString(100))
                End If


            Else
                MsgBox("SO Details and Daily Marketing Schedule uploaded. ", MsgBoxStyle.OkOnly, ResolveResString(100))
            End If

            dtpDateFrom.Enabled = False : dtpDateTo.Enabled = False
            dtpSOdate.Enabled = False : txtCustomerCode.Enabled = False
            txtPlantCode.Enabled = False : txtFileName.Enabled = False
            txtRemarks.Enabled = False : txtECESS.Enabled = False
            txtAddExcise.Enabled = False : txtVat.Enabled = False
            txtExciseDuty.Enabled = False : txtSECESS.Enabled = False
            cmdHelpAddExcise.Enabled = False : cmdHelpEcess.Enabled = False
            cmdHelpSHECess.Enabled = False : cmdVATHelp.Enabled = False
            cmdExDutyHelp.Enabled = False : txtAddExcise.Enabled = False
            txtECESS.Enabled = False : txtExciseDuty.Enabled = False
            txtSECESS.Enabled = False : txtVat.Enabled = False
            ChkCT2reqd.Enabled = False
            cmdSODetails.Text = "Update SO"
            mP_Connection.CommitTrans()
        Else
            MsgBox("SO Details and Daily Marketing Schedules are not uploaded due to some error. ", MsgBoxStyle.Information, ResolveResString(100))
            cmdSODetails.Enabled = True
            cmdSODetails.Text = "Sales Order Details"
            mP_Connection.RollbackTrans()
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        mP_Connection.RollbackTrans()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        frmExport.ShowDialog()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        If gblnCancelExport Then Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) : Exit Sub
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function CheckForMultipleOpenSO(ByVal Account_Code As String, ByVal Cons_Code As String, ByVal Cust_ref As String, ByVal Cust_DrgNo As String, ByVal Item_code As String, ByVal openso As Boolean) As String
        Dim rstHelpDb As ClsResultSetDB
        Dim blnopenclosedso As Boolean

        Try
            rstHelpDb = New ClsResultSetDB

            Call rstHelpDb.GetResult("Select dbo.UDF_CHECK_ACTIVE_SO_ITEM('" & gstrUNITID & "' ,'" & Account_Code & "','" & Cons_Code & "','" & Cust_ref & "','" & Cust_DrgNo & "','" & Item_code & "','" & openso & "') as ActiveSalesOrder", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rstHelpDb.GetNoRows >= 1 Then
                If Len(rstHelpDb.GetValue("ActiveSalesOrder")) > 0 Then
                    CheckForMultipleOpenSO = rstHelpDb.GetValue("ActiveSalesOrder")
                Else
                    CheckForMultipleOpenSO = ""
                End If
            End If
            rstHelpDb.ResultSetClose()
            rstHelpDb = Nothing
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub cmdVATHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdVATHelp.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "VAT Help", 1)
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
    Private Sub dtpDateFrom_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDateFrom.ValueChanged
        On Error GoTo ErrHandler
        If dtpDateFrom.Value > dtpDateTo.Value Then
            dtpDateFrom.Value = dtpDateTo.Value
            dtpSOdate.Value = dtpDateFrom.Value
        End If
        If dtpDateFrom.Value > dtpSOdate.Value Then
            dtpSOdate.Value = dtpDateFrom.Value
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpDateFrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtpDateFrom.KeyPress
        On Error GoTo ErrHandler
        Dim keyAscii As Short = Asc(e.KeyChar)
        If keyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpDateTo_ValueChanged(ByVal eventSender As System.Object, ByVal e As System.EventArgs) Handles dtpDateTo.ValueChanged
        On Error GoTo ErrHandler
        If dtpDateFrom.Value > dtpDateTo.Value Then
            dtpDateTo.Value = dtpDateFrom.Value
            dtpSOdate.Value = dtpDateFrom.Value
        End If
        If dtpSOdate.Value > dtpDateTo.Value Then
            dtpDateTo.Value = dtpSOdate.Value
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpDateTo_KeyPressEvent(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtpDateTo.KeyPress
        On Error GoTo ErrHandler
        Dim keyAscii As Short = Asc(e.KeyChar)
        If keyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpSOdate_ValueChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpSOdate.ValueChanged
        On Error GoTo ErrHandler
        If bool_value_change = True Then
            Exit Sub
        End If
        'If CStr(Year(getDateForDB(dtpSOdate.Value))) + CStr(Month(getDateForDB(dtpSOdate.Value))) > CStr(Year(getDateForDB(dtpDateTo.Value))) + CStr(Month(getDateForDB(dtpDateTo.Value))) Then
        If CDate(Format(dtpSOdate.Value, "yyyy MMM")) > CDate(Format(dtpDateTo.Value, "yyyy MMM")) Then
            bool_value_change = True
            dtpSOdate.Value = dtpDateTo.Value
            bool_value_change = False
        End If
        'If CStr(Year(getDateForDB(dtpSOdate.Value))) + CStr(Month(getDateForDB(dtpSOdate.Value))) < CStr(Year(GetServerDate())) + CStr(Month(GetServerDate())) Then
            'bool_value_change = True
            'dtpSOdate.Value = GetServerDate()
            'bool_value_change = False
        'End If
        'If CStr(Year(getDateForDB(dtpSOdate.Value))) + CStr(Month(getDateForDB(dtpSOdate.Value))) < CStr(Year(getDateForDB(dtpDateFrom.Value))) + CStr(Month(getDateForDB(dtpDateFrom.Value))) Then
        If CDate(Format(dtpSOdate.Value, "yyyy MMM")) > CDate(Format(dtpSOdate.Value, "yyyy MMM")) Then
            bool_value_change = True
            dtpSOdate.Value = dtpDateFrom.Value
            bool_value_change = False
        End If
        Select Case Month(getDateForDB(dtpSOdate.Value))
            Case 1
                lblMonth.Text = "Jan"
            Case 2
                lblMonth.Text = "Feb"
            Case 3
                lblMonth.Text = "Mar"
            Case 4
                lblMonth.Text = "Apr"
            Case 5
                lblMonth.Text = "May"
            Case 6
                lblMonth.Text = "June"
            Case 7
                lblMonth.Text = "July"
            Case 8
                lblMonth.Text = "Aug"
            Case 9
                lblMonth.Text = "Sep"
            Case 10
                lblMonth.Text = "Oct"
            Case 11
                lblMonth.Text = "Nov"
            Case 12
                lblMonth.Text = "Dec"
            Case Else
                lblMonth.Text = ""
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpSOdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtpSOdate.KeyPress
        On Error GoTo ErrHandler
        Dim keyAscii As Short = Asc(e.KeyChar)
        If keyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0057_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        'Form number assigned for the current form
        mdifrmMain.CheckFormName = mintFormIndex
        'Form name text is made BOLD
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0057_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        'Form name would be adjusted to normal
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0057_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader_ClickEvent(ctlFormHeader, New System.EventArgs())
        End If
    End Sub
    Private Sub frmMKTTRN0057_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim rsDtpDateTo As New ClsResultSetDB
        Dim StrSql As String

        On Error GoTo ErrHandler
        'Adjusts the form to the main Window standards.
        Call FitToClient(Me, frmMain, ctlFormHeader, frmButton, 500)
        'Assigns the integer value of the form
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.HeaderString())
        Call EnableControls(True, Me)
        dtpDateFrom.Format = DateTimePickerFormat.Custom
        dtpDateFrom.CustomFormat = gstrDateFormat
        dtpDateFrom.Value = GetServerDate()
        Call ChkVBSchUpdFlag()
        'dtpDateTo.Format = DateTimePickerFormat.Custom
        'dtpDateTo.CustomFormat = gstrDateFormat
        'dtpDateTo.Value = GetServerDate()
        '10725400--Starts
        'If gstrUNITID <> "MPU" Then
        '    dtpDateTo.Format = DateTimePickerFormat.Custom
        '    dtpDateTo.CustomFormat = gstrDateFormat
        '    dtpDateTo.Value = GetServerDate()
        'Else
        '    StrSql = "Select Top 1 Fin_End_date from Financial_Year_Tb where UNIT_CODE= '" & gstrUNITID & "' and GETDATE() between Fin_Start_date and Fin_End_date"
        '    If IsRecordExists(StrSql) = True Then
        '        rsDtpDateTo.GetResult(StrSql)
        '        dtpDateTo.Text = VB6.Format(rsDtpDateTo.GetValue("Fin_End_date"), gstrDateFormat)
        '    Else
        '        dtpDateTo.Text = GetServerDate()
        '    End If
        'End If
        'Anupam Kumar - INC1589893
        Dim FinEnddate() As String = {"MPU", "MC2", "MNK", "MRG"}
        If Array.IndexOf(FinEnddate, gstrUNITID) >= 0 Then
            StrSql = "Select Top 1 Fin_End_date from Financial_Year_Tb where UNIT_CODE= '" & gstrUNITID & "' and GETDATE() between Fin_Start_date and Fin_End_date"
            If IsRecordExists(StrSql) = True Then
                rsDtpDateTo.GetResult(StrSql)
                dtpDateTo.Text = VB6.Format(rsDtpDateTo.GetValue("Fin_End_date"), gstrDateFormat)
            Else
                dtpDateTo.Text = GetServerDate()
            End If

        Else
            dtpDateTo.Format = DateTimePickerFormat.Custom
            dtpDateTo.CustomFormat = gstrDateFormat
            dtpDateTo.Value = GetServerDate()
        End If
        'Anupam Kumar - INC1589893
        rsDtpDateTo.ResultSetClose()
        rsDtpDateTo = Nothing
        '10725400--Ends

        dtpSOdate.Format = DateTimePickerFormat.Custom
        dtpSOdate.CustomFormat = "MMM yyyy"
        dtpSOdate.Value = GetServerDate()

        lblCredit_days.Visible = False
        LblCurrencyCode.Text = ""
        LblCurrencyCode.Visible = False
        cmbSearch.Items.Add("Item Code")
        cmbSearch.Items.Add("CustPart Code")
        cmdPlantHelp.Enabled = False
        txtPlantCode.Enabled = False
        txtPlantCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

        txtDocNo.MaxLength = 18
        mblnDUPLICATE_SOALLOWED = DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE ALLOW_DUPLICATE_SOUPLD = 1 and UNIT_CODE = '" & gstrUNITID & "'")
        '10736222-changes done for CT2 ARE 3 functionality
        ChkCT2reqd.Enabled = False
        Call FN_Spread_Settings()
        ChkVBSchUpdFlag()
        'GST CHANGES
        If gblnGSTUnit = True Then
            cmdHelpAddExcise.Enabled = False
            cmdHelpEcess.Enabled = False
            cmdHelpSHECess.Enabled = False
            cmdVATHelp.Enabled = False
            cmdExDutyHelp.Enabled = False
            txtAddExcise.Enabled = False
            txtECESS.Enabled = False : txtExciseDuty.Enabled = False
            txtSECESS.Enabled = False : txtVat.Enabled = False
            'cmdchangetype.Enabled = False
            txtExciseDuty.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtAddExcise.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtECESS.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtSECESS.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtVat.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
        'GST CHANGES
        If IsRecordExists("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE UPDATE_DMS_SO_SCH_UPLD = 0 and UNIT_CODE = '" & gstrUNITID & "'") Then
            mblnUPDATE_DMS_SO_SCH_UPLD = False
        Else
            mblnUPDATE_DMS_SO_SCH_UPLD = True
        End If

        If mblnUPDATE_DMS_SO_SCH_UPLD = False Then
            CmdUpdSchedule.Enabled = False
        Else
            CmdUpdSchedule.Enabled = True
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0057_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
        rsitemdesc.GetResult("SELECT ITEM_DESC,DRG_DESC FROM CUSTITEM_MST WHERE ITEM_CODE = '" & varInternalPart & "' AND CUST_DRGNO = '" & varCustPart & "' AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND ACTIVE = 1 AND UNIT_CODE='" & gstrUNITID & "'")
        If rsitemdesc.GetNoRows > 0 Then
            lblItemDesc.Text = rsitemdesc.GetValue("ITEM_DESC")
            lblCustPartDesc.Text = rsitemdesc.GetValue("DRG_DESC")
        Else
            lblItemDesc.Text = ""
            lblCustPartDesc.Text = ""
        End If
        DisplayDailyMktSchedule(True, 0)
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SSPurOrd_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SSPurOrd.KeyDownEvent
        On Error GoTo Errorhandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        If e.keyCode = Keys.F1 And e.shift = 0 Then
            With SSPurOrd
                Select Case .ActiveCol
                    Case ENUM_Grid.AddEx
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "Additional Excise Duty")
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                .Col = ENUM_Grid.AddEx : .Row = .ActiveRow
                                .Value = strHelp(0)
                            Else
                                MsgBox(" No record available", vbInformation, ResolveResString(100))
                            End If
                        End If
                    Case ENUM_Grid.Cess
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "ECS", 1)
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                .Col = ENUM_Grid.Cess : .Row = .ActiveRow
                                .Value = strHelp(0)
                            Else
                                MsgBox(" No record available", vbInformation, ResolveResString(100))
                            End If
                        End If
                    Case ENUM_Grid.EDPer
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "Excise Duty", 1)
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                .Col = ENUM_Grid.EDPer : .Row = .ActiveRow
                                .Value = strHelp(0)
                            Else
                                MsgBox(" No record available", vbInformation, ResolveResString(100))
                            End If
                        End If
                    Case ENUM_Grid.SHECess
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "SHECESS", 1)
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                .Col = ENUM_Grid.SHECess : .Row = .ActiveRow
                                .Value = strHelp(0)
                            Else
                                MsgBox(" No record available", vbInformation, ResolveResString(100))
                            End If
                        End If
                    Case ENUM_Grid.VAT
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "'", "VALUE ADDED TAX", 1)
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                .Col = ENUM_Grid.VAT : .Row = .ActiveRow
                                .Value = strHelp(0)
                            Else
                                MsgBox(" No record available", vbInformation, ResolveResString(100))
                            End If
                        End If
                End Select
            End With
        End If
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SSPurOrd_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SSPurOrd.KeyPressEvent
        On Error GoTo ErrHandler
        If SSPurOrd.ActiveCol = ENUM_Grid.custsupply Or SSPurOrd.ActiveCol = ENUM_Grid.ToolCost Then
            If e.keyAscii = 45 Or e.keyAscii = 43 Then
                e.keyAscii = 0
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SSPurOrd_LeaveCell1(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SSPurOrd.LeaveCell
        On Error GoTo ErrHandler
        Dim rsTax As New ClsResultSetDB
        With SSPurOrd
            If e.col >= ENUM_Grid.EDPer And e.col <= ENUM_Grid.VAT Then
                .Col = e.col
                .Row = e.row
                If Len(LTrim(RTrim(.Value))) > 0 Then
                    rsTax.GetResult("select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where txrt_rate_no = '" & .Value & "' and UNIT_CODE='" & gstrUNITID & "'")
                    If rsTax.GetNoRows <= 0 Then
                        MsgBox("Invalid Tax.", vbInformation + vbOKOnly, ResolveResString(100))
                        .Value = ""
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                        Exit Sub
                    End If
                End If
            End If
            If e.newRow >= 1 And e.newRow <= SSPurOrd.MaxRows Then
                Call DisplayDailyMktSchedule(False, e.newRow)
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAddExcise_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAddExcise.KeyPress
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
    Private Sub txtAddExcise_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAddExcise.Leave
        On Error GoTo ErrHandler
        Dim rsTax As New ClsResultSetDB
        If Len(LTrim(RTrim(txtAddExcise.Text))) > 0 Then
            rsTax.GetResult("select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where txrt_rate_no = '" & txtAddExcise.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
            If rsTax.GetNoRows <= 0 Then
                MsgBox("Invalid Tax.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtAddExcise.Text = ""
                txtAddExcise.Focus()
            End If
            rsTax.ResultSetClose()
        End If
        Exit Sub
ErrHandler:
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
    Private Sub txtCustomerCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.Leave
        On Error GoTo ErrHandler
        Dim rsCust As New ClsResultSetDB
        'Changes against 10737738 
        If SchUpdFlag = True Then
            rsCust.GetResult("Select customer_Code from customer_mst where customer_code = '" & txtCustomerCode.Text & "' and UNIT_CODE='" & gstrUNITID & "' and SCH_UPLOAD_CODE ='SOUPLOAD' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
        Else
            rsCust.GetResult("Select customer_Code from customer_mst where customer_code = '" & txtCustomerCode.Text & "' and UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
        End If
        If Len(LTrim(RTrim(txtCustomerCode.Text))) > 0 Then
            If rsCust.GetNoRows <= 0 Then
                MsgBox("Invalid Customer.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                txtCustomerCode.Text = ""
                txtCustomerCode.Focus()
            End If
        End If
        rsCust.ResultSetClose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim bolExist As Boolean
        Dim StrSQL As String = String.Empty
        If Len(txtCustomerCode.Text) > 0 Then
            bolExist = ValCustomerCode()
            blnlinelevelcustomer = Find_Value("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")
            blnSOUPLD_NotRemovehypen = Find_Value("SELECT SOUPLD_NotRemovehypen FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")
            mblnSOUPLD_ForexCurrency = Find_Value("SELECT SOUPLD_FOREXCURRENCY FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'")

            
        End If
        If bolExist = True Then
            Call display()
        End If

        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        If Len(txtCustomerCode.Text.Trim) > 0 Then
            strSQL = "Select dbo.UDF_IsCT2Customer('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
            ChkCT2reqd.Enabled = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(StrSQL))
        Else
            ChkCT2reqd.Enabled = False
            ChkCT2reqd.Checked = False
        End If

        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtDocNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocNo.TextChanged
        On Error GoTo ErrHandler
        Dim strDocNo As String
        strDocNo = txtDocNo.Text
        Flag = True
        If Len(Trim(txtDocNo.Text)) = 0 Then
            Call RefreshForm()
        End If
        Flag = False
        txtDocNo.Text = strDocNo
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDocNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Flag = True
            Call RefreshForm()
            Flag = False
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
        lblItemDesc.Text = ""
        lblCustPartDesc.Text = ""
        Call Initialize_SpdSOSch()
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
    Private Sub txtDocNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDocNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim docno As Integer
        Dim StrSql As String
        Dim rsDocNo As New ClsResultSetDB
        Dim RSNot As New ClsResultSetDB
        If Len(Trim(txtDocNo.Text)) > 0 Then
            StrSql = "SELECT distinct H.DOC_NO,M.CUSTOMER_CODE,M.Cust_Name,D.VALID_FROM,  D.VALID_TO FROM SO_UPLD_dtl D INNER JOIN SO_UPLD_HDR H  ON D.DOC_NO = H.DOC_NO AND D.UNIT_CODE = H.UNIT_CODE AND H.DOC_NO NOT IN (SELECT DISTINCT ISNULL(DOC_NO,0)  FROM dailymkTschedule WHERE FILETYPE = 'SOUPLD' AND UNIT_CODE='" & gstrUNITID & "') AND H.DOC_NO = '" & txtDocNo.Text & "'  INNER JOIN CUSTOMER_MST M ON  M.CUSTOMER_CODE = H.CUST_CODE "
            StrSql = StrSql & " AND M.UNIT_CODE = H.UNIT_CODE WHERE D.UNIT_CODE='" & gstrUNITID & "' and ((isnull(M.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= M.deactive_date))"
            rsDocNo.GetResult(StrSql)
            RSNot.GetResult("SELECT DOC_NO FROM SO_UPLD_HDR WHERE DOC_NO = '" & txtDocNo.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
            If RSNot.GetNoRows <= 0 Then
                MsgBox("No Record Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Call RefreshForm()
                RSNot.ResultSetClose()
                GoTo EventExitSub
            Else
                If rsDocNo.GetNoRows > 0 Then
                    Call FetchSORecords(docno)
                    cmdSODetails.Enabled = True
                    cmdSODetails.Text = "Update SO"
                    CmdUpdSchedule.Enabled = True
                Else
                    Call FetchSORecords(docno)
                    cmdSODetails.Enabled = False
                    CmdUpdSchedule.Enabled = False
                End If
            End If
            RSNot.ResultSetClose()
            rsDocNo.ResultSetClose()
            If SSPurOrd.MaxRows > 0 Then
                SSPurOrd.Col = ENUM_Grid.custsupply
                SSPurOrd.Row = 1
                SSPurOrd.Focus()
            End If
        Else
            Call RefreshForm()
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
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtECESS_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtECESS.Leave
        On Error GoTo ErrHandler
        Dim rsTax As New ClsResultSetDB
        If Len(LTrim(RTrim(txtECESS.Text))) > 0 Then
            rsTax.GetResult("select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where txrt_rate_no = '" & txtECESS.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
            If rsTax.GetNoRows <= 0 Then
                MsgBox("Invalid Tax.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtECESS.Text = ""
                txtECESS.Focus()
            End If
            rsTax.ResultSetClose()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
    Private Sub txtExciseDuty_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExciseDuty.Leave
        On Error GoTo ErrHandler
        Dim rsTax As New ClsResultSetDB
        If Len(LTrim(RTrim(txtExciseDuty.Text))) > 0 Then
            rsTax.GetResult("select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where txrt_rate_no = '" & txtExciseDuty.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
            If rsTax.GetNoRows <= 0 Then
                MsgBox("Invalid Tax.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtExciseDuty.Text = ""
                txtExciseDuty.Focus()
            End If
            rsTax.ResultSetClose()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
    Private Sub txtSECESS_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSECESS.Leave
        On Error GoTo ErrHandler
        Dim rsTax As New ClsResultSetDB
        If Len(LTrim(RTrim(txtSECESS.Text))) > 0 Then
            rsTax.GetResult("select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where txrt_rate_no = '" & txtSECESS.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
            If rsTax.GetNoRows <= 0 Then
                MsgBox("Invalid Tax.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtSECESS.Text = ""
                txtSECESS.Focus()
            End If
            rsTax.ResultSetClose()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
            If Not CheckRecord("select distinct C.wh_code as PlantCode,W.WH_Description as PlantName  from custwarehouse_mst C, warehouse_mst W  where C.customer_code = '" & txtCustomerCode.Text & "'  and c.wh_code = w.wh_code and c.UNIT_CODE = w.UNIT_CODE AND W.WH_CODE = '" & txtPlantCode.Text & "' and c.UNIT_CODE='" & gstrUNITID & "'") Then
                MsgBox(" Invalid Plant Code", MsgBoxStyle.Information, ResolveResString(100))
                Me.txtPlantCode.Text = "" : Me.txtPlantCode.Focus()
                Cancel = True
            End If
        End If
        If Cancel = True Then
            Me.txtPlantCode.Focus()
        Else
            Call rsobject.GetResult("Select WH_Description from warehouse_mst  where wh_code = '" & Trim(Me.txtPlantCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "'")
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
    Private Function SO_Upload() As Object
        On Error GoTo Errorhandler
        Dim Obj_EX As New Excel.Application
        Dim Row As Short
        Dim Row_Ex As Short
        Dim Col_Ex As Short
        Dim rsItemCode As New ClsResultSetDB
        Dim RSCUSTSUPPLY As New ClsResultSetDB
        Dim MissingItem As String
        Dim RSDUPLICATE As New ClsResultSetDB
        Dim STRDUPLICATE As String = ""
        Dim strTariffMissing As String = ""  '10530629
        Dim strresult As String = String.Empty
        Dim STRHSNCODE As String = String.Empty
        Dim MissingDiscount As String = ""
        Dim mblnDiscountMandatory As Boolean
        mblnDiscountMandatory = CBool(Find_Value("SELECT isnull(SOUPLD_Discounteditable,0)   FROM customer_mst WHERE  customer_code='" & txtCustomerCode.Text.Trim() & "' and UNIT_CODE='" + gstrUNITID + "'"))
        If cmdSODetails.Text = "Update SO" Then
            update_SO()
            Return Nothing
            Exit Function
        End If

        Row_Ex = 2
        Row = 1
        SSPurOrd.MaxRows = 1
        Obj_EX.Workbooks.Open(Trim(txtFileName.Text))
        
        If txtECESS.Text = "" Then
            txtECESS.Text = ""
        End If
        If txtSECESS.Text = "" Then
            txtSECESS.Text = ""
        End If
        If txtVat.Text = "" Then
            txtVat.Text = ""
        End If
        MissingItem = ""
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        STRDUPLICATE = ""

        strresult = Obj_EX.Range("$A$" & Row_Ex).Value2.ToString.Trim
        While strresult <> ""

            rsItemCode.GetResult("SELECT DISTINCT ITEM_CODE,item_desc,DRG_DESC FROM CUSTITEM_MST" & _
                                 " WHERE CUST_DRGNO = '" & Obj_EX.Range("$E$" & Row_Ex).Value2 & "' AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND ACTIVE = 1 AND UNIT_CODE='" & gstrUNITID & "'")
            '" WHERE CUST_DRGNO = '" & Replace(Obj_EX.Range("$E$" & Row_Ex).Value2, "-", "") & "' AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND ACTIVE = 1 AND UNIT_CODE='" & gstrUNITID & "'")
            If rsItemCode.GetNoRows <= 0 Then
                MissingItem = MissingItem & Obj_EX.Range("$E$" & Row_Ex).Value2 & vbCrLf
                'MissingItem = MissingItem & Replace(Obj_EX.Range("$E$" & Row_Ex).Value2, "-", "") & vbCrLf
            End If
            If blnSOUPLD_NOTRemovehypen = True Then
                RSDUPLICATE.GetResult("select DISTINCT C.ACCOUNT_CODE, C.cust_drgno,COUNT(C.item_code)" & _
                " countitem from custitem_mst C with (nolock) where C.active = 1 " & _
                " and CUST_DRGNO = '" & Obj_EX.Range("$E$" & Row_Ex).Value2 & "' " & _
                " AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' and UNIT_CODE = '" & gstrUNITID & "' GROUP BY ACCOUNT_CODE, CUST_DRGNO")
                
            Else
                RSDUPLICATE.GetResult("select DISTINCT C.ACCOUNT_CODE, C.cust_drgno,COUNT(C.item_code)" & _
                " countitem from custitem_mst C with (nolock) where C.active = 1 " & _
                " and CUST_DRGNO = '" & Replace(Obj_EX.Range("$E$" & Row_Ex).Value2, "-", "") & "' " & _
                " AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' and UNIT_CODE = '" & gstrUNITID & "' GROUP BY ACCOUNT_CODE, CUST_DRGNO")
            End If
            If RSDUPLICATE.GetNoRows > 0 Then
                If RSDUPLICATE.GetValue("COUNTITEM") > 1 Then
                    If STRDUPLICATE.Length = 0 Then
                        STRDUPLICATE = Replace(Obj_EX.Range("$E$" & Row_Ex).Value2, "-", "")
                    Else
                        STRDUPLICATE = STRDUPLICATE + "," + Replace(Obj_EX.Range("$E$" & Row_Ex).Value2, "-", "")
                    End If
                Else
                    '10530629 
                    If Not IsRecordExists("select top 1 1 from item_mst where UNIT_CODE = '" & gstrUNITID & "'" & _
                            " and Item_Code = '" & rsItemCode.GetValue("ITEM_CODE") & "'" & _
                            " and ISNULL(tariff_code,'') <> ''") And gblnGSTUnit = False Then
                        If strTariffMissing.Length = 0 Then
                            strTariffMissing = rsItemCode.GetValue("ITEM_CODE")
                        Else
                            strTariffMissing = strTariffMissing + "," + rsItemCode.GetValue("ITEM_CODE")
                        End If
                    End If
                End If
            End If
            With SSPurOrd
                Row = .MaxRows
                .SetText(ENUM_Grid.salesorder, .MaxRows, Obj_EX.Range("$B$" & Row_Ex).Value2.ToString)
                If blnSOUPLD_NOTRemovehypen = True Then
                    .SetText(ENUM_Grid.CustPartNo, .MaxRows, Obj_EX.Range("$E$" & Row_Ex).Value2)
                Else
                    .SetText(ENUM_Grid.CustPartNo, .MaxRows, Replace(Obj_EX.Range("$E$" & Row_Ex).Value2, "-", ""))
                End If

                If rsItemCode.GetNoRows > 0 Then
                    .SetText(ENUM_Grid.InternalPartNo, .MaxRows, rsItemCode.GetValue("item_code"))
                    .SetText(ENUM_Grid.ItemDesc, .MaxRows, rsItemCode.GetValue("item_desc"))
                    .SetText(ENUM_Grid.CustPartDesc, .MaxRows, rsItemCode.GetValue("DRG_desc"))
                Else
                    .SetText(ENUM_Grid.InternalPartNo, .MaxRows, "")
                    .SetText(ENUM_Grid.ItemDesc, .MaxRows, "")
                    .SetText(ENUM_Grid.CustPartDesc, .MaxRows, "")
                End If
                .SetText(ENUM_Grid.ShopCode, .MaxRows, Obj_EX.Range("$D$" & Row_Ex).Value2)
                .SetText(ENUM_Grid.UOM, .MaxRows, Obj_EX.Range("$H$" & Row_Ex).Value2)
                If txtPlantCode.Text.Trim = "" Then
                    txtPlantCode.Text = Obj_EX.Range("$c$" & Row_Ex).Value2
                End If
                RSCUSTSUPPLY = New ClsResultSetDB
                RSCUSTSUPPLY.GetResult("SELECT CUST_MTRL,TOOL_COST FROM CUSTITEM_MST  WHERE ACCOUNT_CODE = '" & txtCustomerCode.Text & "'  AND ITEM_CODE = '" & rsItemCode.GetValue("ITEM_CODE") & "'   AND ACTIVE = 1 AND UNIT_CODE='" & gstrUNITID & "'")
                .SetText(ENUM_Grid.Qty, .MaxRows, Replace(Obj_EX.Range("$G$" & Row_Ex).Value2, ",", ""))
                If UCase(Trim(gstrUNITID)) = "MPU" Or UCase(Trim(gstrUNITID)) = "SMC" Or UCase(Trim(gstrUNITID)) = "SMP" Or UCase(Trim(gstrUNITID)) = "MRG" Or UCase(Trim(gstrUNITID)) = "VF1" Or UCase(Trim(gstrUNITID)) = "MGS" Or UCase(Trim(gstrUNITID)) = "VF2" Or UCase(Trim(gstrUNITID)) = "SA2" Or UCase(Trim(gstrUNITID)) = "DUR" Or UCase(Trim(gstrUNITID)) = "MNK" Or UCase(Trim(gstrUNITID)) = "MC2" Then
                    .SetText(ENUM_Grid.ExWorks, .MaxRows, Obj_EX.Range("$J$" & Row_Ex).Value2.ToString())
                    .SetText(ENUM_Grid.ToolCost, .MaxRows, Obj_EX.Range("$I$" & Row_Ex).Value2)
                ElseIf UCase(Trim(gstrUNITID)) = "MST" Or UCase(Trim(gstrUNITID)) = "MSB" Or UCase(Trim(gstrUNITID)) = "MAR" Or UCase(Trim(gstrUNITID)) = "SDA" Then
                    .SetText(ENUM_Grid.ExWorks, .MaxRows, Obj_EX.Range("$J$" & Row_Ex).Value2.ToString())
                    .SetText(ENUM_Grid.ToolCost, .MaxRows, Obj_EX.Range("$K$" & Row_Ex).Value2)
                Else
                    .SetText(ENUM_Grid.ExWorks, .MaxRows, (Obj_EX.Range("$I$" & Row_Ex).Value2 / Obj_EX.Range("$J$" & Row_Ex).Value2).ToString())
                End If

                If Not (UCase(Trim(gstrUNITID)) = "MPU" Or UCase(Trim(gstrUNITID)) = "SMC" Or UCase(Trim(gstrUNITID)) = "SMP" Or UCase(Trim(gstrUNITID)) = "MST" Or UCase(Trim(gstrUNITID)) = "SDA" Or UCase(Trim(gstrUNITID)) = "MSB") Or UCase(Trim(gstrUNITID)) = "MNK" Then
                    If RSCUSTSUPPLY.GetNoRows > 0 Then
                        .SetText(ENUM_Grid.custsupply, .MaxRows, IIf(Len(Trim(CStr(RSCUSTSUPPLY.GetValue("CUST_MTRL")))) = 0, "0", RSCUSTSUPPLY.GetValue("CUST_MTRL")))
                        .SetText(ENUM_Grid.ToolCost, .MaxRows, IIf(Len(Trim(CStr(RSCUSTSUPPLY.GetValue("TOOL_COST")))) = 0, "0.00", RSCUSTSUPPLY.GetValue("TOOL_COST")))
                    Else
                        .SetText(ENUM_Grid.custsupply, .MaxRows, "0")
                        .SetText(ENUM_Grid.ToolCost, .MaxRows, Obj_EX.Range("$I$" & Row_Ex).Value2)
                    End If
                End If
                RSCUSTSUPPLY.ResultSetClose()
                Dim strDiscounttype As String = ""
                If UCase(Trim(gstrUNITID)) = "MST" Or UCase(Trim(gstrUNITID)) = "MSB" Or UCase(Trim(gstrUNITID)) = "SDA" Then
                    .SetText(ENUM_Grid.discount_type, .MaxRows, Obj_EX.Range("$R$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.discount_value, .MaxRows, Obj_EX.Range("$S$" & Row_Ex).Value2)
                    strDiscounttype = Obj_EX.Range("$R$" & Row_Ex).Value2
                Else
                    .SetText(ENUM_Grid.discount_type, .MaxRows, Obj_EX.Range("$Q$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.discount_value, .MaxRows, Obj_EX.Range("$R$" & Row_Ex).Value2)
                    strDiscounttype = Obj_EX.Range("$Q$" & Row_Ex).Value2
                End If
                ' added by priti
                If mblnDiscountMandatory Then
                    If (UCase(strDiscounttype) <> "V" And UCase(strDiscounttype) <> "P") Then
                        MissingDiscount = MissingDiscount & Obj_EX.Range("$E$" & Row_Ex).Value2 & vbCrLf
                    End If
                End If


                .SetText(ENUM_Grid.EDPer, .MaxRows, LTrim(RTrim(txtExciseDuty.Text)))
                .SetText(ENUM_Grid.AddEx, .MaxRows, LTrim(RTrim(txtAddExcise.Text)))
                .SetText(ENUM_Grid.Cess, .MaxRows, txtECESS.Text)
                .SetText(ENUM_Grid.SHECess, .MaxRows, txtSECESS.Text)
                .SetText(ENUM_Grid.VAT, .MaxRows, txtVat.Text)
                .SetText(ENUM_Grid.Remarks, .MaxRows, txtRemarks.Text)
                'gst changes
                If UCase(Trim(gstrUNITID)) = "MPU" Or UCase(Trim(gstrUNITID)) = "MC2" Or UCase(Trim(gstrUNITID)) = "SMC" Or UCase(Trim(gstrUNITID)) = "SMP" Or UCase(Trim(gstrUNITID)) = "MRG" Or UCase(Trim(gstrUNITID)) = "MNK" Then
                    .SetText(ENUM_Grid.CGST_TAX, .MaxRows, Obj_EX.Range("$k$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.SGST_TAX, .MaxRows, Obj_EX.Range("$L$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.IGST_TAX, .MaxRows, Obj_EX.Range("$M$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.COMP_ECC, .MaxRows, Obj_EX.Range("$N$" & Row_Ex).Value2)
                    '101528442
                ElseIf UCase(Trim(gstrUNITID)) = "MST" Or UCase(Trim(gstrUNITID)) = "MSB" Or UCase(Trim(gstrUNITID)) = "MAR" Or UCase(Trim(gstrUNITID)) = "SDA" Then
                    .SetText(ENUM_Grid.Packing_Per, .MaxRows, Obj_EX.Range("$L$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.CGST_TAX, .MaxRows, Obj_EX.Range("$M$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.SGST_TAX, .MaxRows, Obj_EX.Range("$N$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.IGST_TAX, .MaxRows, Obj_EX.Range("$O$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.COMP_ECC, .MaxRows, Obj_EX.Range("$P$" & Row_Ex).Value2)
                Else
                    '' This else part is using for M03 Hyundai 
                    .SetText(ENUM_Grid.CGST_TAX, .MaxRows, Obj_EX.Range("$k$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.IGST_TAX, .MaxRows, Obj_EX.Range("$L$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.SGST_TAX, .MaxRows, Obj_EX.Range("$M$" & Row_Ex).Value2)
                    .SetText(ENUM_Grid.COMP_ECC, .MaxRows, Obj_EX.Range("$N$" & Row_Ex).Value2)
                End If
                STRHSNCODE = Find_Value("SELECT HSN_SAC_CODE FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & rsItemCode.GetValue("ITEM_CODE") & "'")
                .SetText(ENUM_Grid.HSN_CODE, .MaxRows, STRHSNCODE)
                'gst changes
            End With
            Row_Ex = Row_Ex + 1
            SSPurOrd.MaxRows = SSPurOrd.MaxRows + 1

            'Issue Id           :10114399
            If Obj_EX.Range("$A$" & Row_Ex).Value2 Is Nothing = True Then GoTo ExitLoop
            strresult = Obj_EX.Range("$A$" & Row_Ex).Value2.ToString
        End While

ExitLoop:
        SSPurOrd.MaxRows = SSPurOrd.MaxRows - 1
        Obj_EX.Workbooks.Close()
        If Len(LTrim(RTrim(MissingItem))) <> 0 Then
            MsgBox("These Items are Not Defined In The System." & vbCrLf & vbCrLf & MissingItem & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            SSPurOrd.MaxRows = 0
            Return Nothing
            Exit Function
        End If
        If Len(LTrim(RTrim(STRDUPLICATE))) <> 0 Then
            MsgBox("These Customer Drawing Nos are Active For Multiple Items In Customer Item Master." & vbCrLf & vbCrLf & STRDUPLICATE & vbCrLf & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            SSPurOrd.MaxRows = 0
            Return Nothing
            Exit Function
        End If

        '10530629 
        If Len(LTrim(RTrim(strTariffMissing))) <> 0 Then
            MsgBox("For Following items Tariff Code not defined in Item Master." & vbCrLf & vbCrLf & strTariffMissing & vbCrLf & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            SSPurOrd.MaxRows = 0
            Return Nothing
            Exit Function
        End If

        If Len(LTrim(RTrim(MissingDiscount))) <> 0 Then
            MsgBox("For Following items Discount type not defined in Excel." & vbCrLf & vbCrLf & MissingDiscount & vbCrLf & vbCrLf & "Can't Upload File...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            SSPurOrd.MaxRows = 0
            Return Nothing
            Exit Function
        End If


        Call Upload_SO()


        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Return Nothing
        Exit Function
Errorhandler:
        If Err.Number = 1004 Then
            MsgBox("File Not Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            SSPurOrd.MaxRows = 0
            Exit Function
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function RefreshForm() As Object
        On Error GoTo Errorhandler
        Dim Row, Col As Short
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        txtCustomerCode.Text = ""
        txtPlantCode.Text = ""
        txtFileName.Text = ""
        SSPurOrd.MaxRows = 0
        dtpDateFrom.Value = GetServerDate()
        dtpDateTo.Value = GetServerDate()
        dtpSOdate.Value = GetServerDate()
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
        txtRemarks.Enabled = True
        txtRemarks.Text = ""
        cmdSODetails.Enabled = True
        cmdSODetails.Text = "Sales Order Details"
        txtCustomerCode.Enabled = True
        txtPlantCode.Enabled = True
        txtFileName.Enabled = True
        cmdCustHelp.Enabled = True
        cmdPlantHelp.Enabled = True
        cmdFileHelp.Enabled = True
        dtpDateFrom.Enabled = True
        dtpDateFrom.Value = GetServerDate()
        dtpDateTo.Enabled = True
        dtpDateTo.Value = GetServerDate()
        dtpSOdate.Enabled = True
        dtpSOdate.Value = GetServerDate()
        lblCredit_days.Text = ""
        LblCurrencyCode.Text = ""
        LblCurrencyCode.Visible = False
        lblCustPartDesc.Text = ""
        lblItemDesc.Text = ""
        txtShipAddress.Text = ""
        lblShipAddress.Text = ""
        lblShipAddress_Details.Text = ""
        chkShipAddress.Checked = False
        If mblnUPDATE_DMS_SO_SCH_UPLD = True Then
            CmdUpdSchedule.Enabled = True
        End If
        txtSearch.Text = ""
        If gblnGSTUnit = False Then
            cmdHelpAddExcise.Enabled = True
            cmdHelpEcess.Enabled = True
            cmdHelpSHECess.Enabled = True
            cmdVATHelp.Enabled = True
            cmdExDutyHelp.Enabled = True
            txtAddExcise.Enabled = True
            txtECESS.Enabled = True : txtExciseDuty.Enabled = True
            txtSECESS.Enabled = True : txtVat.Enabled = True
            cmdchangetype.Enabled = True
        Else
            cmdHelpAddExcise.Enabled = False
            cmdHelpEcess.Enabled = False
            cmdHelpSHECess.Enabled = False
            cmdVATHelp.Enabled = False
            cmdExDutyHelp.Enabled = False
            txtAddExcise.Enabled = False
            txtECESS.Enabled = False : txtExciseDuty.Enabled = False
            txtSECESS.Enabled = False : txtVat.Enabled = False
            '            cmdchangetype.Enabled = False
            txtExciseDuty.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtAddExcise.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtECESS.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtSECESS.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtVat.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If

        '10736222-changes done for CT2 ARE 3 functionality
        ChkCT2reqd.Enabled = False
        ChkCT2reqd.Checked = False
        Call FN_Spread_Settings()
        For Row = 1 To 4
            For Col = 1 To 10
                SpdSOSch.SetText(Col, Row, "")
            Next
        Next
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Return Nothing
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
        Call mobjEmpDll.CRecordset.OpenRecordset("select * from customer_mst where customer_code = '" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
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
            mobjEmpDll.CRecordset.OpenRecordset("select * from customer_mst where Customer_code = '" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rsLocation.GetResult("select * from customer_mst where Customer_code = '" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "'")
            rsLocation.ResultSetClose()
            If Not mobjEmpDll.CRecordset.EOF_Renamed Then
                mobjEmpDll.CRecordset.MoveFirst()
                txtCustomerCode.Text = mobjEmpDll.CRecordset.GetFieldValue("customer_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                LblCustomerName.Text = mobjEmpDll.CRecordset.GetFieldValue("cust_name", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                lblCredit_days.Text = mobjEmpDll.CRecordset.GetFieldValue("Credit_days", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                LblCurrencyCode.Text = mobjEmpDll.CRecordset.GetFieldValue("currency_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
            End If
        End If
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
        Dim rsDespQty As ADODB.Recordset
        Dim varDespatchQty As Object = Nothing
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call FN_Spread_Settings()
        dtpDateFrom.Enabled = False : dtpDateTo.Enabled = False
        dtpSOdate.Enabled = False : txtCustomerCode.Enabled = False
        txtPlantCode.Enabled = False : txtFileName.Enabled = False
        'txtRemarks.Enabled = False : 
        cmdCustHelp.Enabled = False
        cmdPlantHelp.Enabled = False : cmdFileHelp.Enabled = False
        cmdHelpAddExcise.Enabled = False
        cmdHelpEcess.Enabled = False
        cmdHelpSHECess.Enabled = False
        cmdVATHelp.Enabled = False
        cmdExDutyHelp.Enabled = False
        txtAddExcise.Enabled = False
        txtECESS.Enabled = False : txtExciseDuty.Enabled = False
        txtSECESS.Enabled = False : txtVat.Enabled = False
        cmdchangetype.Enabled = False
        Row = 1
        StrSql = "SELECT dbo.SO_Upld_dtl.*,M.CUST_NAME,h.plant_c,h.file_location,  h.remarks AS REMARKS_HDR ,W.WH_Description   FROM dbo.SO_Upld_hdr h INNER JOIN dbo.SO_Upld_dtl  ON h.Doc_no = dbo.SO_Upld_dtl.Doc_no AND h.UNIT_CODE = dbo.SO_Upld_dtl.UNIT_CODE AND  h.cust_code = dbo.SO_Upld_dtl.CUST_CODE  INNER JOIN CUSTOMER_MST M ON M.CUSTOMER_CODE = H.CUST_CODE AND M.UNIT_CODE = H.UNIT_CODE  INNER JOIN WAREHOUSE_MST W ON W.WH_CODE = H.PLANT_C AND"
        StrSql = StrSql & " W.UNIT_CODE = H.UNIT_CODE AND h.DOC_NO = '" & txtDocNo.Text & "' WHERE H.UNIT_CODE='" & gstrUNITID & "' order by SalesOrder"
        rsRecords = New ClsResultSetDB
        rsRecords.GetResult(StrSql)
        If rsRecords.RowCount > 0 Then
            rsRecords.MoveFirst()
            txtCustomerCode.Text = rsRecords.GetValue("cust_code")
            txtPlantCode.Text = rsRecords.GetValue("plant_c")
            txtFileName.Text = rsRecords.GetValue("file_location")
            dtpDateFrom.Value = rsRecords.GetValue("valid_from")
            dtpDateTo.Value = rsRecords.GetValue("valid_to")
            dtpSOdate.Value = rsRecords.GetValue("So_Date")
            txtExciseDuty.Text = rsRecords.GetValue("EDPER")
            txtAddExcise.Text = rsRecords.GetValue("SHECESS")
            txtVat.Text = rsRecords.GetValue("VAT")
            txtECESS.Text = rsRecords.GetValue("CESS")
            txtSECESS.Text = rsRecords.GetValue("SHECESS")
            txtRemarks.Text = rsRecords.GetValue("REMARKS_HDR")
            LblCustomerName.Text = rsRecords.GetValue("CUST_NAME")
            lblPlantName.Text = rsRecords.GetValue("WH_DESCRIPTION")
            While Not rsRecords.EOFRecord
                With SSPurOrd
                    SSPurOrd.MaxRows = SSPurOrd.MaxRows + 1
                    .SetText(ENUM_Grid.salesorder, .MaxRows, rsRecords.GetValue("SalesOrder"))
                    .SetText(ENUM_Grid.CustPartNo, .MaxRows, rsRecords.GetValue("PartNo"))
                    .SetText(ENUM_Grid.InternalPartNo, .MaxRows, rsRecords.GetValue("item_code"))
                    .SetText(ENUM_Grid.ItemDesc, .MaxRows, rsRecords.GetValue("PartName"))
                    .SetText(ENUM_Grid.CustPartDesc, .MaxRows, rsRecords.GetValue("SalesOrder"))
                    .SetText(ENUM_Grid.Qty, .MaxRows, rsRecords.GetValue("Quantity"))
                    .SetText(ENUM_Grid.ExWorks, .MaxRows, rsRecords.GetValue("ExWorks"))
                    .SetText(ENUM_Grid.custsupply, .MaxRows, IIf(Len(Trim(CStr(rsRecords.GetValue("CustSupply")))) = 0, "0", rsRecords.GetValue("CustSupply")))
                    .SetText(ENUM_Grid.ToolCost, .MaxRows, IIf(Len(Trim(CStr(rsRecords.GetValue("ToolCost")))) = 0, "0.00", rsRecords.GetValue("ToolCost")))
                    .SetText(ENUM_Grid.EDPer, .MaxRows, rsRecords.GetValue("EDPer"))
                    .SetText(ENUM_Grid.AddEx, .MaxRows, rsRecords.GetValue("AddEx"))
                    .SetText(ENUM_Grid.Cess, .MaxRows, rsRecords.GetValue("Cess"))
                    .SetText(ENUM_Grid.SHECess, .MaxRows, rsRecords.GetValue("SHECess"))
                    .SetText(ENUM_Grid.VAT, .MaxRows, rsRecords.GetValue("VAT"))
                    .SetText(ENUM_Grid.Remarks, .MaxRows, rsRecords.GetValue("Remarks"))
                    .SetText(ENUM_Grid.PrevQty, .MaxRows, rsRecords.GetValue("PrevQty"))
                    .SetText(ENUM_Grid.ShopCode, .MaxRows, rsRecords.GetValue("shop_code"))
                    .SetText(ENUM_Grid.UOM, .MaxRows, rsRecords.GetValue("uom"))
                    .SetText(ENUM_Grid.discount_type, .MaxRows, rsRecords.GetValue("discount_type"))
                    .SetText(ENUM_Grid.discount_value, .MaxRows, rsRecords.GetValue("discount_value"))


                    rsDespQty = New ADODB.Recordset
                    rsDespQty.Open("select isnull(despatchqty,0) despatchqty FROM so_upld_dtl where PartNo = '" & rsRecords.GetValue("PartNo") & "' and cust_code = '" & txtCustomerCode.Text & "' and DOC_NO = " & txtDocNo.Text & " and UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
                    If Not rsDespQty.EOF And Not rsDespQty.BOF Then
                        varDespatchQty = Nothing
                            varDespatchQty = CInt(rsDespQty.Fields("DESPATCHQTY").Value)
                        .Col = ENUM_Grid.DespatchQty : .Row = .MaxRows
                        .Text = varDespatchQty
                    Else
                        .Col = ENUM_Grid.DespatchQty : .Row = .MaxRows
                        .Text = "0"
                    End If
                End With
                rsRecords.MoveNext()
            End While
            rsWkngDays.GetResult("select distinct work_flg,dt from  calendar_mfg_mst where month (dt) = '" & Month(getDateForDB(dtpSOdate.Value)) & "'  and YEAR (dt) = '" & Year(getDateForDB(dtpSOdate.Value)) & "' and UNIT_CODE='" & gstrUNITID & "' order by dt")
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
            If SpdSOSch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) Then
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
        Else
            FetchSORecords = "No Record Found"
        End If
        rsRecords.ResultSetClose()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Return Nothing
        Exit Function
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub Initialize_SpdSOSch()
        Dim Col, Row As Short
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        For Row = 1 To 4
            For Col = 1 To 10
                SpdSOSch.SetText(Col, Row, "")
            Next
        Next
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
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
    Function CheckUploaded_Data() As Boolean
        Dim StrSql As String
        Dim rsRecords As ClsResultSetDB
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        StrSql = "SELECT DBO.SO_UPLD_HDR.DOC_NO"
        StrSql = StrSql & " FROM DBO.SO_UPLD_HDR "
        StrSql = StrSql & " WHERE DBO.SO_UPLD_HDR.CUST_CODE = '" & Trim(txtCustomerCode.Text) & "'"
        StrSql = StrSql & " AND DBO.SO_UPLD_HDR.PLANT_C = '" & Trim(txtPlantCode.Text) & "'"
        StrSql = StrSql & " AND DBO.SO_UPLD_HDR.FILE_LOCATION = '" & Replace(Trim(txtFileName.Text), "'", "''") & "'"
        StrSql = StrSql & " AND DBO.SO_UPLD_HDR.UNIT_CODE = '" & gstrUNITID & "'"
        rsRecords = New ClsResultSetDB
        rsRecords.GetResult(StrSql)
        If rsRecords.RowCount > 0 Then
            CheckUploaded_Data = True
            Me.txtDocNo.Text = rsRecords.GetValue("DOC_NO")
        Else
            CheckUploaded_Data = False
        End If
        rsRecords.ResultSetClose()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Function DeleteExistsData() As Boolean
        Dim StrSql As String
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        StrSql = "DELETE FROM DBO.SO_UPLD_HDR WHERE DBO.SO_UPLD_HDR.DOC_NO = '" & Trim(txtDocNo.Text) & "' AND DBO.SO_UPLD_HDR.UNIT_CODE='" & gstrUNITID & "'"
        StrSql = StrSql & " DELETE FROM DBO.SO_UPLD_DTL WHERE DBO.SO_UPLD_DTL.DOC_NO = '" & Trim(txtDocNo.Text) & "' AND  DBO.SO_UPLD_DTL.UNIT_CODE='" & gstrUNITID & "'"
        StrSql = StrSql & " DELETE FROM DBO.DAILYMKTSCHEDULE_TEMP WHERE DBO.DAILYMKTSCHEDULE_TEMP.DOC_NO = '" & Trim(txtDocNo.Text) & "' AND DBO.DAILYMKTSCHEDULE_TEMP.FILETYPE = 'SOUPLD' AND DBO.DAILYMKTSCHEDULE_TEMP.UNIT_CODE='" & gstrUNITID & "'"
        mP_Connection.Execute(StrSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        DeleteExistsData = True
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        DeleteExistsData = False
    End Function
    Private Function update_SO() As Object
        On Error GoTo ErrHandler
        Dim varCustSupply As Object = Nothing
        Dim varToolCost As Object = Nothing
        Dim varpackingper As Object = Nothing
        Dim varED As Object = Nothing
        Dim varAddED As Object = Nothing
        Dim varCess As Object = Nothing
        Dim varSHECess As Object = Nothing
        Dim varVAT As Object = Nothing
        Dim varRemarks As Object = Nothing
        Dim varSalesOrder As Object = Nothing
        Dim varItemCode As Object = Nothing
        Dim strReturnCustRef As String

        If cmdSODetails.Text = "Update SO" Then
            With SSPurOrd
                varCustSupply = Nothing
                varToolCost = Nothing
                varpackingper = Nothing
                varED = Nothing
                varAddED = Nothing
                varCess = Nothing
                varSHECess = Nothing
                varVAT = Nothing
                varRemarks = Nothing
                varSalesOrder = Nothing
                varItemCode = Nothing
                SSPurOrd.GetText(ENUM_Grid.custsupply, SSPurOrd.ActiveRow, varCustSupply)
                SSPurOrd.GetText(ENUM_Grid.ToolCost, SSPurOrd.ActiveRow, varToolCost)
                SSPurOrd.GetText(ENUM_Grid.Packing_Per, SSPurOrd.ActiveRow, varpackingper)
                SSPurOrd.GetText(ENUM_Grid.EDPer, SSPurOrd.ActiveRow, varED)
                SSPurOrd.GetText(ENUM_Grid.AddEx, SSPurOrd.ActiveRow, varAddED)
                SSPurOrd.GetText(ENUM_Grid.Cess, SSPurOrd.ActiveRow, varCess)
                SSPurOrd.GetText(ENUM_Grid.SHECess, SSPurOrd.ActiveRow, varSHECess)
                SSPurOrd.GetText(ENUM_Grid.VAT, SSPurOrd.ActiveRow, varVAT)
                SSPurOrd.GetText(ENUM_Grid.Remarks, SSPurOrd.ActiveRow, varRemarks)
                SSPurOrd.GetText(ENUM_Grid.salesorder, SSPurOrd.ActiveRow, varSalesOrder)
                SSPurOrd.GetText(ENUM_Grid.InternalPartNo, .ActiveRow, varItemCode)
            End With
        End If
        If LTrim(RTrim(txtDocNo.Text)) <> "" Then
            mP_Connection.Execute("update so_upld_dtl set packing_per='" & varpackingper & "', custsupply = '" & varCustSupply & "' ,  toolcost = " & IIf(Len(LTrim(RTrim(CStr(varToolCost)))) = 0, 0, varToolCost) & ",EDPer = '" & varED & "',  SHECess = '" & varSHECess & "',Cess = '" & varCess & "',  VAT = '" & varVAT & "',AddEx = '" & varAddED & "',REMARKS = '" & Replace(varRemarks, "'", "''") & "'  where cust_code = '" & txtCustomerCode.Text & "' and  salesorder = '" & varSalesOrder & "' and doc_no = " & txtDocNo.Text & " and  item_code = '" & varItemCode & "' and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            MsgBox("Record Updated Successfully Against Item : " & varItemCode, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
        Else
            MsgBox("Invalid Doc No.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
        End If
        Return Nothing
        Exit Function
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Function
    Private Sub txtVat_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVat.Leave
        On Error GoTo ErrHandler
        Dim rsTax As New ClsResultSetDB
        If Len(LTrim(RTrim(txtVat.Text))) > 0 Then
            rsTax.GetResult("select txrt_rate_no,txrt_ratedesc from Gen_TaxRate where txrt_rate_no = '" & txtVat.Text & "' and UNIT_CODE='" & gstrUNITID & "'")
            If rsTax.GetNoRows <= 0 Then
                MsgBox("Invalid Tax.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtVat.Text = ""
                txtVat.Focus()
            End If
        End If
        rsTax.ResultSetClose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
    '10797956
    Private Sub ChkCT2reqd_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkCT2reqd.CheckedChanged
        Try
            If ChkCT2reqd.Checked = True Then
                txtExciseDuty.Text = "EX0"
                txtAddExcise.Text = ""
                txtECESS.Text = "EC0"
                txtSECESS.Text = "ECSSH0"
                txtAddExcise.Enabled = False
                txtExciseDuty.Enabled = False
                txtECESS.Enabled = False
                txtSECESS.Enabled = False
                cmdExDutyHelp.Enabled = False
                cmdHelpEcess.Enabled = False
                cmdHelpAddExcise.Enabled = False
                cmdHelpSHECess.Enabled = False
            Else
                If gblnGSTUnit = False Then
                    txtExciseDuty.Text = ""
                    txtAddExcise.Text = ""
                    txtECESS.Text = ""
                    txtSECESS.Text = ""
                    txtAddExcise.Enabled = True
                    txtExciseDuty.Enabled = True
                    txtECESS.Enabled = True
                    txtSECESS.Enabled = True
                    cmdExDutyHelp.Enabled = True
                    cmdHelpEcess.Enabled = True
                    cmdHelpAddExcise.Enabled = True
                    cmdHelpSHECess.Enabled = True
                End If
            End If
        Catch ex As Exception

        End Try
        
    End Sub
    Private Sub chkShipAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShipAddress.CheckedChanged
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            If chkShipAddress.Checked = True Then
                StrSql = "select Distinct Shipping_Code,Shipping_Desc,Ship_Address1,Ship_Address2,Ship_State,GSTIN_ID from Customer_Shipping_Dtl where unit_code='" & gstrUNITID & "' and InActive_Flag=0 and customer_code='" & Trim(txtCustomerCode.Text) & "'"
                StrSrvCHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Ship Address Code Help")
                If UBound(StrSrvCHelp) <= 0 Then
                    chkShipAddress.Checked = False
                    Exit Sub
                End If

                If StrSrvCHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtShipAddress.Text = "" : txtShipAddress.Focus() : Exit Sub
                Else
                    txtShipAddress.Text = StrSrvCHelp(0)
                    '                            lblservicedesc.Text = StrSrvCHelp(1)
                    lblShipAddress_Details.Text = StrSrvCHelp(1)
                End If

            Else
                txtShipAddress.Text = ""
                lblShipAddress_Details.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

End Class