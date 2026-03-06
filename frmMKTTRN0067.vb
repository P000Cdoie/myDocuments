'Developed By   :   Siddharth Ranjan
'Revised By     :   Shabbir Hussain
'Revision Date  :   15 Nov 2010
'                   Added another option for entering Item Category wise sales value
'Modified By Sanchi on 23 May 2011
'   Modified to support MultiUnit functionality
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmMKTTRN0067
    Dim mintFormIndex As Short
    Dim mIs_Freezed As Boolean
    Dim mstrCategory As String = ""
    Enum enmGrid3
        col_Cust_code = 1
        col_help = 2
        col_Cust_name = 3
        col_Year_Sale = 4
        col_Total_sale = 5
        col_Original_Sales_of_day = 6
        col_Sales_of_day = 7
        col_Type = 8
        col_adjust = 9
        col_adj_hidden = 10
        col_hidden2 = 11
    End Enum
    Enum enmGrid1
        col_Unit = 1
        col_Sale_for_Day = 2
        col_MTD_Budget = 3
        col_MTD_actual = 4
        col_MVariance = 5
        col_Month_Budget = 6
        col_YTD_Budget = 7
        col_YTD_Actual = 8
        col_YVariance = 9
        col_Year_Budget = 10
        col_current_rate_day = 11
        col_Ask_rate_day = 12
    End Enum
#Region "Item Category Wise Adjustment Grid"
    Private Enum enumsspr
        Item_Category = 1
        Customer_Code
        Hlp
        Customer_Name
        Sale_For_Day
        Month_Budget
        Month_Actual
        Month_Variance
        YTM_Budget
        YTM_Actual
        YTM_Variance
        Adjustment_Type
        Sale_Value
        Current_Rate_Day
        Ask_Rate_Day
    End Enum
    Private Sub SetSpreadProperty()
        Dim i As Integer
        Try
            With ssprCategoryWiseAdjustment
                .MaxRows = 0
                .MaxCols = 0
                .MaxCols = enumsspr.Ask_Rate_Day
                .ColHeaderRows = 2
                .Row = 0
                .Col = enumsspr.Item_Category : .Text = "Category" : .set_ColWidth(enumsspr.Item_Category, 800)
                .Col = enumsspr.Customer_Code : .Text = "Customer Code" : .set_ColWidth(enumsspr.Customer_Code, 800)
                .Col = enumsspr.Hlp : .Text = "Hlp" : .set_ColWidth(enumsspr.Hlp, 300)
                .Col = enumsspr.Customer_Name : .Text = "Customer Name" : .set_ColWidth(enumsspr.Customer_Name, 2800) : .ColHidden = True
                .Col = enumsspr.Sale_For_Day : .Text = "Sales For The Day" : .set_ColWidth(enumsspr.Sale_For_Day, 1000)
                .Col = enumsspr.Month_Budget : .Text = "Month" : .set_ColWidth(enumsspr.Month_Budget, 1000)
                .Col = enumsspr.Month_Actual : .Text = "Month" : .set_ColWidth(enumsspr.Month_Actual, 1000)
                .Col = enumsspr.Month_Variance : .Text = "Variance" : .set_ColWidth(enumsspr.Month_Variance, 1000)
                .Col = enumsspr.YTM_Budget : .Text = "YTM" : .set_ColWidth(enumsspr.YTM_Budget, 1000)
                .Col = enumsspr.YTM_Actual : .Text = "YTM" : .set_ColWidth(enumsspr.YTM_Actual, 1000)
                .Col = enumsspr.YTM_Variance : .Text = "Variance" : .set_ColWidth(enumsspr.YTM_Variance, 900)
                .Col = enumsspr.Adjustment_Type : .Text = "Type" : .set_ColWidth(enumsspr.Adjustment_Type, 800)
                .Col = enumsspr.Sale_Value : .Text = "Sale Value" : .set_ColWidth(enumsspr.Sale_Value, 1400)
                .Col = enumsspr.Current_Rate_Day : .Text = "Current Rate Per Day" : .set_ColWidth(enumsspr.Current_Rate_Day, 1200) : .ColHidden = True
                .Col = enumsspr.Ask_Rate_Day : .Text = "Ask. Rate Per Day" : .set_ColWidth(enumsspr.Ask_Rate_Day, 1200) : .ColHidden = True
                .set_RowHeight(0, 270)
                .Row = 1
                .Col = enumsspr.Item_Category : .Text = "Category" : .set_ColWidth(enumsspr.Item_Category, 800)
                .Col = enumsspr.Customer_Code : .Text = "Customer Code" : .set_ColWidth(enumsspr.Customer_Code, 800)
                .Col = enumsspr.Hlp : .Text = "Hlp" : .set_ColWidth(enumsspr.Hlp, 300)
                .Col = enumsspr.Customer_Name : .Text = "Customer Name" : .set_ColWidth(enumsspr.Customer_Name, 2800) : .ColHidden = True
                .Col = enumsspr.Sale_For_Day : .Text = "Sales For The Day" : .set_ColWidth(enumsspr.Sale_For_Day, 1000)
                .Col = enumsspr.Month_Budget : .Text = "Budget" : .set_ColWidth(enumsspr.Month_Budget, 1000)
                .Col = enumsspr.Month_Actual : .Text = "Actual" : .set_ColWidth(enumsspr.Month_Actual, 1000)
                .Col = enumsspr.Month_Variance : .Text = "Variance" : .set_ColWidth(enumsspr.Month_Variance, 1000)
                .Col = enumsspr.YTM_Budget : .Text = "Budget" : .set_ColWidth(enumsspr.YTM_Budget, 1000)
                .Col = enumsspr.YTM_Actual : .Text = "Actual" : .set_ColWidth(enumsspr.YTM_Actual, 1000)
                .Col = enumsspr.YTM_Variance : .Text = "Variance" : .set_ColWidth(enumsspr.YTM_Variance, 900)
                .Col = enumsspr.Adjustment_Type : .Text = "Type" : .set_ColWidth(enumsspr.Adjustment_Type, 800)
                .Col = enumsspr.Sale_Value : .Text = "Sale Value" : .set_ColWidth(enumsspr.Sale_Value, 1400)
                .Col = enumsspr.Current_Rate_Day : .Text = "Current Rate Per Day" : .set_ColWidth(enumsspr.Current_Rate_Day, 1200) : .ColHidden = True
                .Col = enumsspr.Ask_Rate_Day : .Text = "Ask. Rate Per Day" : .set_ColWidth(enumsspr.Ask_Rate_Day, 1200) : .ColHidden = True
                .set_RowHeight(1, 270)
                .Enabled = True
                .Row = 0
                For i = enumsspr.Item_Category To enumsspr.Sale_For_Day
                    .Col = i
                    .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted
                    .RowMerge = FPSpreadADO.MergeConstants.MergeRestricted
                Next i
                .Col = enumsspr.Month_Variance
                .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted
                .RowMerge = FPSpreadADO.MergeConstants.MergeRestricted
                For i = enumsspr.YTM_Variance To enumsspr.Ask_Rate_Day
                    .Col = i
                    .ColMerge = FPSpreadADO.MergeConstants.MergeRestricted
                    .RowMerge = FPSpreadADO.MergeConstants.MergeRestricted
                Next i
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub SetSpreadColTypes(ByVal pintRowNo As Integer)
        Try
            With ssprCategoryWiseAdjustment
                .Row = pintRowNo
                .Col = enumsspr.Item_Category : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Customer_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Hlp : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeButtonPicture = My.Resources.ico111.ToBitmap
                .Col = enumsspr.Customer_Name : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Sale_For_Day : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Month_Budget : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Month_Actual : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Month_Variance : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.YTM_Budget : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.YTM_Actual : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.YTM_Variance : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Adjustment_Type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .TypeComboBoxList = "Dr.(+)" + Chr(9) + "Cr.(-)"
                .Col = enumsspr.Sale_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = "99,999,999,999,999" : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = False
                .Col = enumsspr.Sale_Value : .EditModePermanent = True : .EditModeReplace = True
                .Col = enumsspr.Current_Rate_Day : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Ask_Rate_Day : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub addRowAtEnterKeyPress(ByVal pintRows As Integer)
        Dim intRowHeight As Integer
        Try
            With ssprCategoryWiseAdjustment
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                For intRowHeight = 1 To pintRows
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .set_RowHeight(.Row, 330)
                    Call SetSpreadColTypes(.Row)
                Next intRowHeight
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
#End Region
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdclose.Click
        Me.Close()
    End Sub
    Private Sub frmMKTTRN0067_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0067_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0067_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0067_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        On Error GoTo ErrHandler
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlHeader.Tag)
        Call FillLabelFromResFile(Me)
        Call FitToClient(Me, fraContainer, ctlHeader, GrpButtons, 550)
        Call EnableControls(False, Me, True)
        SetSpreadProperty()
        addRowAtEnterKeyPress(1)
        MdiParent = prjMPower.mdifrmMain
        CmdDelete.Enabled = False
        initialize_Grid1()
        initialize_Grid3()
        SetSpreadProperty()
        populateDate_Data("Load")
        PopulateItemCategory()
        cmdclose.Enabled = True
        cmdSave.Enabled = False
        cmdFreeze.Enabled = False
        cmdExport.Enabled = True
        AddBlankRow.Enabled = True
        cmdShowData.Enabled = True
        ProgressBar1.Enabled = True
        cmdTrackerHlp.Enabled = True
        cmdRefresh.Enabled = True
        cmdGenerateBudget.Enabled = True
        cmdPrint.Enabled = True
        cmdReportTracker.Enabled = True
        txtReportTracker.Enabled = True
        cmdAliasReport.Enabled = True
        txtReportTracker.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Grid3.Enabled = True
        txtSearch.Enabled = True
        txtSearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdSearch.Enabled = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub initialize_Grid3()
        With Grid3
            .MaxRows = 0
            .MaxCols = 11
            .set_RowHeight(0, 20)
            .SetText(enmGrid3.col_Cust_code, .MaxRows, "Customer Code")
            .SetText(enmGrid3.col_help, .MaxRows, "Hlp")
            .SetText(enmGrid3.col_Cust_name, .MaxRows, "Customer Name")
            .SetText(enmGrid3.col_Year_Sale, .MaxRows, "Year Sale")
            .SetText(enmGrid3.col_Total_sale, .MaxRows, "Month Sales Amt.")
            .SetText(enmGrid3.col_Original_Sales_of_day, .MaxRows, "Sales For The Day (Original)")
            .SetText(enmGrid3.col_Sales_of_day, .MaxRows, "Sales For The Day (Adjusted)")
            .SetText(enmGrid3.col_Type, .MaxRows, "Type")
            .SetText(enmGrid3.col_adjust, .MaxRows, "Adj. Amt.")
            .SetText(enmGrid3.col_adj_hidden, .MaxRows, "Adj.Hidden")
            .SetText(enmGrid3.col_hidden2, .MaxRows, "Hidden2")
            .set_ColWidth(enmGrid3.col_Cust_code, 10)
            .set_ColWidth(enmGrid3.col_help, 4)
            .set_ColWidth(enmGrid3.col_Cust_name, 35)
            .set_ColWidth(enmGrid3.col_Year_Sale, 10)
            .set_ColWidth(enmGrid3.col_Total_sale, 10)
            .set_ColWidth(enmGrid3.col_Original_Sales_of_day, 12)
            .set_ColWidth(enmGrid3.col_Sales_of_day, 12)
            .set_ColWidth(enmGrid3.col_Type, 7)
            .set_ColWidth(enmGrid3.col_adjust, 13)
            .set_ColWidth(enmGrid3.col_adj_hidden, 13)
            .set_ColWidth(enmGrid3.col_hidden2, 13)
            .Col = enmGrid3.col_Year_Sale
            .ColHidden = True
            .Col = enmGrid3.col_Total_sale
            .ColHidden = True
            .Col = enmGrid3.col_adj_hidden
            .ColHidden = True
            .Col = enmGrid3.col_hidden2
            .ColHidden = True
        End With
    End Sub
    Private Sub initialize_Grid1()
        With fpsGrid1
            .MaxRows = 0
            .MaxCols = 12
            .set_RowHeight(1, 30)
            .SetText(enmGrid1.col_Unit, 0, "Unit")
            .SetText(enmGrid1.col_Unit, 1, "Unit")
            .SetText(enmGrid1.col_Sale_for_Day, 0, "Sales For The Day")
            .SetText(enmGrid1.col_Sale_for_Day, 1, "Sales For The Day")
            .SetText(enmGrid1.col_MTD_Budget, 0, "MTD")
            .SetText(enmGrid1.col_MTD_actual, 0, "MTD")
            .SetText(enmGrid1.col_MTD_Budget, 1, "Budget")
            .SetText(enmGrid1.col_MTD_actual, 1, "Actual")
            .SetText(enmGrid1.col_MVariance, 0, "Variance")
            .SetText(enmGrid1.col_MVariance, 1, "Variance")
            .SetText(enmGrid1.col_Month_Budget, 0, "Month Budget")
            .SetText(enmGrid1.col_Month_Budget, 1, "Month Budget")
            .SetText(enmGrid1.col_YTD_Actual, 0, "YTD")
            .SetText(enmGrid1.col_YTD_Budget, 0, "YTD")
            .SetText(enmGrid1.col_YTD_Budget, 1, "Budget")
            .SetText(enmGrid1.col_YTD_Actual, 1, "Actual")
            .SetText(enmGrid1.col_YVariance, 0, "Variance")
            .SetText(enmGrid1.col_YVariance, 1, "Variance")
            .SetText(enmGrid1.col_Year_Budget, 0, "Year Budget")
            .SetText(enmGrid1.col_Year_Budget, 1, "Year Budget")
            .SetText(enmGrid1.col_current_rate_day, 0, "Current Rate Per Day")
            .SetText(enmGrid1.col_current_rate_day, 1, "Current Rate Per Day")
            .SetText(enmGrid1.col_Ask_rate_day, 0, "Asking Rate Per Day")
            .SetText(enmGrid1.col_Ask_rate_day, 1, "Asking Rate Per Day")
            .set_ColWidth(enmGrid1.col_current_rate_day, 9)
            .Col = enmGrid1.col_Sale_for_Day
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_MTD_Budget
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_MTD_actual
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_MVariance
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_YTD_Budget
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_YTD_Actual
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_YVariance
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_Month_Budget
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_Year_Budget
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_current_rate_day
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_Ask_rate_day
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid1.col_Unit
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_Sale_for_Day
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_MVariance
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_MTD_Budget
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_Month_Budget
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_YVariance
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_Year_Budget
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_current_rate_day
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
            .Col = enmGrid1.col_Ask_rate_day
            .Row = 0
            .ColMerge = FPSpreadADO.MergeConstants.MergeAlways
            .RowMerge = FPSpreadADO.MergeConstants.MergeAlways
        End With
    End Sub
    Private Sub Add_Grid3_blankRow()
        With Grid3
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.MaxRows, 12)
            .Col = enmGrid3.col_Total_sale
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
            .Col = enmGrid3.col_Sales_of_day
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
            .Col = enmGrid3.col_Original_Sales_of_day
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
            .Col = enmGrid3.col_Type
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
            .TypeComboBoxList = "Dr.(+)" + Chr(9) + "Cr.(-)"
            '.TypeComboBoxCurSel = 0
            .Col = enmGrid3.col_adjust
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatMax = "99,999,999,999,999"
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid3.col_adj_hidden
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatMax = "99,999,999,999,999"
            .TypeFloatDecimalPlaces = 4
            .Col = enmGrid3.col_hidden2
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatMax = "99,999,999,999,999"
            .TypeFloatDecimalPlaces = 4
            .Value = 0.0
            .Col = enmGrid3.col_help
            '.Row = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .TypeButtonPicture = My.Resources.ico111.ToBitmap
            .Col = enmGrid3.col_Cust_code
            .Col2 = enmGrid3.col_Cust_code
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            .Col = enmGrid3.col_Cust_name
            .Col2 = enmGrid3.col_Sales_of_day
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
    End Sub
    Private Sub Add_Grid1_blankRow()
        With fpsGrid1
            .MaxRows = .MaxRows + 1
            .set_RowHeight(.MaxRows, 22)
            .Col = enmGrid1.col_Unit
            .Col2 = enmGrid1.col_Ask_rate_day
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
    End Sub
    Private Sub populate_Grid1(ByVal str_Tracker_No As String)
        Dim objconn As SqlConnection = Nothing
        Dim objDR As SqlDataReader
        Dim objCommand As New SqlCommand()
        Try
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = "SELECT UNIT,round(SUM(SALES_FOR_DAY),8), round(SUM(MTD_BUDGET),8), round(SUM(MTD_ACTUAL),8) ,ROUND(SUM(MONTH_VARIANCE),8), ROUND(SUM(MONTH_BUDGET),8), ROUND(SUM(YTD_BUDGET),8), ROUND(SUM(YTD_ACTUAL),8), ROUND(SUM(YEAR_VARIANCE),8) ,ROUND(SUM(YEAR_BUDGET),8), ROUND(SUM(CURR_RATE_PER_DAY),8), ROUND(SUM(ASKING_RATE_PER_DAY),8) FROM CUST_WISE_SALES_BUDGET WHERE ACTIVE_FLAG =1 AND TRACKER_NO = " & str_Tracker_No & " AND UNIT_CODE='" & gstrUNITID & "' GROUP BY UNIT"
            objCommand.CommandType = CommandType.Text
            objDR = objCommand.ExecuteReader()
            fpsGrid1.MaxRows = 0
            Grid3.MaxRows = 0
            ssprCategoryWiseAdjustment.MaxRows = 0
            While objDR.Read
                Add_Grid1_blankRow()
                fpsGrid1.SetText(enmGrid1.col_Unit, fpsGrid1.MaxRows, objDR.GetValue(0))
                fpsGrid1.SetText(enmGrid1.col_Sale_for_Day, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(1)), 0, objDR.GetValue(1).ToString))
                fpsGrid1.SetText(enmGrid1.col_MTD_Budget, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(2)), 0, objDR.GetValue(2).ToString))
                fpsGrid1.SetText(enmGrid1.col_MTD_actual, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(3)), 0, objDR.GetValue(3).ToString))
                fpsGrid1.SetText(enmGrid1.col_MVariance, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(4)), 0, objDR.GetValue(4).ToString))
                fpsGrid1.SetText(enmGrid1.col_Month_Budget, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(5)), 0, objDR.GetValue(5).ToString))
                fpsGrid1.SetText(enmGrid1.col_YTD_Budget, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(6)), 0, objDR.GetValue(6).ToString))
                fpsGrid1.SetText(enmGrid1.col_YTD_Actual, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(7)), 0, objDR.GetValue(7).ToString))
                fpsGrid1.SetText(enmGrid1.col_YVariance, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(8)), 0, objDR.GetValue(8).ToString))
                fpsGrid1.SetText(enmGrid1.col_Year_Budget, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(9)), 0, objDR.GetValue(9).ToString))
                fpsGrid1.SetText(enmGrid1.col_current_rate_day, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(10)), 0, objDR.GetValue(10).ToString))
                fpsGrid1.SetText(enmGrid1.col_Ask_rate_day, fpsGrid1.MaxRows, IIf(IsDBNull(objDR.GetValue(11)), 0, objDR.GetValue(11).ToString))
            End While
            objDR.Close()
            objDR = Nothing
            objCommand = Nothing
            objconn.Close()
            objconn = Nothing
        Catch ex As Exception
            objDR = Nothing
            objCommand = Nothing
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objconn = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim objconn As SqlConnection = Nothing
        Dim objTrans As SqlTransaction = Nothing
        Dim objCom As New SqlCommand()
        Dim intloop As Int16
        Dim strCustomer_code As Object = ""
        Dim strMonth_sale As Object = Nothing, strYear_sale As Object = Nothing, strSale_of_day As Object = Nothing
        Dim strtype As Object = ""
        Dim stramt As Object = Nothing
        Dim STR As String
        Dim strHidden2 As String = String.Empty
        Dim strItemCategory As Object = Nothing
        Dim dblMonth_Budget As Object = Nothing, dblMonth_Actual As Object = Nothing
        Dim dblYTM_Budget As Object = Nothing, dblYTM_Actual As Object = Nothing
        Dim dblSaleValue As Object = Nothing
        Dim dblCurrentRatePerDay As Double = 0
        Dim dblAskingRatePerDay As Double = 0
        Try
            If validate_Before_save() Then
                objconn = SqlConnectionclass.GetConnection()
                objTrans = objconn.BeginTransaction
                objCom.CommandType = CommandType.Text
                objCom.Connection = objconn
                objCom.Transaction = objTrans
                With Grid3
                    For intloop = 1 To .MaxRows
                        strCustomer_code = Nothing
                        .GetText(enmGrid3.col_Cust_code, intloop, strCustomer_code)
                        strYear_sale = Nothing
                        .GetText(enmGrid3.col_Year_Sale, intloop, strYear_sale)
                        strMonth_sale = Nothing
                        .GetText(enmGrid3.col_Total_sale, intloop, strMonth_sale)
                        strSale_of_day = Nothing
                        .GetText(enmGrid3.col_Sales_of_day, intloop, strSale_of_day)
                        strtype = Nothing
                        .GetText(enmGrid3.col_Type, intloop, strtype)
                        stramt = Nothing
                        .GetText(enmGrid3.col_adjust, intloop, stramt)
                        strHidden2 = Nothing
                        .GetText(enmGrid3.col_hidden2, intloop, strHidden2)
                        If (Convert.ToDouble(stramt) = 0) Then
                            STR = "IF EXISTS(SELECT TOP 1 1 FROM SALES_TRACKING_CUSTOMER_ADJUSTMENT WHERE TRACKER_NO = " & txtTrackerNo.Text & " AND CUSTOMER_CODE = '" & strCustomer_code & "' AND UNIT_CODE='" & gstrUNITID & "')" & _
                                   " DELETE FROM SALES_TRACKING_CUSTOMER_ADJUSTMENT WHERE TRACKER_NO = " & txtTrackerNo.Text & " AND CUSTOMER_CODE = '" & strCustomer_code & "' AND UNIT_CODE='" & gstrUNITID & "'"
                            objCom.CommandText = STR
                            objCom.ExecuteNonQuery()
                        End If
                        If (Convert.ToDouble(stramt) > 0) _
                           And (Convert.ToDouble(stramt) <> Convert.ToDouble(strHidden2)) Then
                            STR = "IF EXISTS(SELECT TOP 1 1 FROM SALES_TRACKING_CUSTOMER_ADJUSTMENT WHERE TRACKER_NO = " & txtTrackerNo.Text & " AND CUSTOMER_CODE = '" & strCustomer_code & "' AND UNIT_CODE='" & gstrUNITID & "')" & _
                                   " DELETE FROM SALES_TRACKING_CUSTOMER_ADJUSTMENT WHERE TRACKER_NO = " & txtTrackerNo.Text & " AND CUSTOMER_CODE = '" & strCustomer_code & "' AND UNIT_CODE='" & gstrUNITID & "'"
                            objCom.CommandText = STR
                            objCom.ExecuteNonQuery()
                            STR = " INSERT INTO SALES_TRACKING_CUSTOMER_ADJUSTMENT(TRACKER_NO,CUSTOMER_CODE,YEAR_SALE,MONTH_SALES,SALE_OF_THE_DAY,ADJ_TYPE,ADJ_AMT,UNIT_CODE)" & _
                                  " VALUES (" & txtTrackerNo.Text & ",'" & strCustomer_code & "'," & strYear_sale & "," & strMonth_sale & "," & strSale_of_day & ",'" & strtype & "'," & Convert.ToDouble(stramt) & ",'" & gstrUNITID & "')"
                            objCom.CommandText = STR
                            objCom.ExecuteNonQuery()
                            If Not UPDATE_CUST_WISE_SALES_BUDGET(strCustomer_code, Mid(strtype, 5, 1), Convert.ToDouble(strMonth_sale), Convert.ToDouble(strYear_sale), Convert.ToDouble(strSale_of_day), (Convert.ToDouble(stramt) / 1000000), objconn, objTrans) Then
                                objTrans.Rollback()
                                objconn.Close()
                                objconn = Nothing
                                objTrans = Nothing
                                Exit Sub
                            End If
                        End If
                        If (Convert.ToDouble(strHidden2) > 0) _
                           And (Convert.ToDouble(stramt) <> Convert.ToDouble(strHidden2)) Then
                            If Mid(strtype, 5, 1) = "+" Then
                                If Not UPDATE_CUST_WISE_SALES_BUDGET(strCustomer_code, "-", Convert.ToDouble(strMonth_sale), Convert.ToDouble(strYear_sale), Convert.ToDouble(strSale_of_day), (Convert.ToDouble(strHidden2) / 1000000), objconn, objTrans) Then
                                    objTrans.Rollback()
                                    objconn.Close()
                                    objconn = Nothing
                                    objTrans = Nothing
                                    Exit Sub
                                End If
                            End If
                            If Mid(strtype, 5, 1) = "-" Then
                                If Not UPDATE_CUST_WISE_SALES_BUDGET(strCustomer_code, "+", Convert.ToDouble(strMonth_sale), Convert.ToDouble(strYear_sale), Convert.ToDouble(strSale_of_day), (Convert.ToDouble(strHidden2) / 1000000), objconn, objTrans) Then
                                    objTrans.Rollback()
                                    objconn.Close()
                                    objconn = Nothing
                                    objTrans = Nothing
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End With
                '------------------------------------------------------------------------------------
                'Added By Shabbir Hussain On Nov 2010
                '------------------------------------------------------------------------------------
                objCom.CommandText = "DELETE FROM ITEMCATEGORY_WISE_CUST_SALES_ADJST WHERE TRACKER_NO = " & txtTrackerNo.Text.Trim & " AND UNIT_CODE='" & gstrUNITID & "'"
                objCom.ExecuteNonQuery()
                If ssprCategoryWiseAdjustment.MaxRows > 0 Then
                    With ssprCategoryWiseAdjustment
                        For intloop = 1 To .MaxRows
                            strItemCategory = Nothing
                            .GetText(enumsspr.Item_Category, intloop, strItemCategory)
                            strCustomer_code = Nothing
                            .GetText(enumsspr.Customer_Code, intloop, strCustomer_code)
                            If strCustomer_code.ToString().Trim.Length > 0 Then
                                Dim objDR As SqlDataReader
                                Dim intAliasCode As Integer = 0, strAliasName As String = ""
                                Dim strCustomerName As String = "", strCustType As String = ""
                                objCom.CommandText = "SELECT  ISNULL(a.ALIAS_CODE,0) as ALIAS_CODE,ISNULL(a.ALIAS_NAME,'') as ALIAS_NAME,B.CUSTOMER_CODE,REPLACE(b.CUST_NAME,'''','') as CUST_NAME ,CUST_TYPE FROM CUSTOMER_MST b  Left Outer join SALES_TRACKING_ALIAS_MAPPING a On(a.UNIT_CODE=b.UNIT_CODE AND a.customer_Code=b.customer_Code and GETDATE() BETWEEN a.EFF_FROM_DATE AND a.EFF_TO_DATE) WHERE B.CUSTOMER_CODE='" & strCustomer_code & "' AND b.UNIT_CODE='" & gstrUNITID & "'"
                                objCom.CommandType = CommandType.Text
                                objDR = objCom.ExecuteReader()
                                While objDR.Read
                                    intAliasCode = Convert.ToInt32(objDR("Alias_Code").ToString())
                                    strAliasName = objDR("Alias_Name").ToString().Trim
                                    strCustomerName = objDR("Cust_Name").ToString().Trim
                                    strCustType = objDR("Cust_Type").ToString().Trim
                                End While
                                If objDR.IsClosed = False Then objDR.Close()
                                objDR = Nothing
                                .Row = intloop
                                .Col = enumsspr.Sale_For_Day
                                strSale_of_day = Val(.Text.Trim.ToString())
                                .Row = intloop
                                .Col = enumsspr.Month_Budget
                                dblMonth_Budget = Val(.Text.Trim.ToString())
                                .Row = intloop
                                .Col = enumsspr.Month_Actual
                                dblMonth_Actual = Val(.Text.Trim.ToString())
                                .Row = intloop
                                .Col = enumsspr.YTM_Budget
                                dblYTM_Budget = Val(.Text.Trim.ToString())
                                .Row = intloop
                                .Col = enumsspr.YTM_Actual
                                dblYTM_Actual = Val(.Text.Trim.ToString())
                                .Row = intloop
                                .Col = enumsspr.Adjustment_Type
                                strtype = .Text.Trim.ToString()
                                .Row = intloop
                                .Col = enumsspr.Sale_Value
                                dblSaleValue = Val(.Text.Trim.ToString())
                                If Val(lblYTD_Work_Day.Text) > 0 Then
                                    dblCurrentRatePerDay = Val(dblYTM_Actual) / Val(lblYTD_Work_Day.Text)
                                End If
                                If (Val(lblWork_Day_Value.Text) - Val(lblYTD_Work_Day.Text)) > 0 Then
                                    dblAskingRatePerDay = (Val(dblYTM_Budget) - Val(dblYTM_Actual)) / (Val(lblWork_Day_Value.Text) - Val(lblYTD_Work_Day.Text))
                                End If
                                STR = " INSERT INTO ITEMCATEGORY_WISE_CUST_SALES_ADJST(TRACKER_NO,ITEM_CATEGORY,ALIAS_CODE,ALIAS_NAME,CUSTOMER_CODE,CUSTOMER_NAME,CUST_TYPE,SALES_FOR_DAY,MONTH_BUDGET,MONTH_ACTUAL,MONTH_VARIANCE,YTM_BUDGET,YTM_ACTUAL,YTM_VARIANCE,ADJ_TYPE,SALES_VALUE,CURR_RATE_PER_DAY,ASKING_RATE_PER_DAY,UNIT_CODE)" & _
                                   " VALUES (" & txtTrackerNo.Text.Trim & ",'" & strItemCategory & "'," & intAliasCode & ",'" & strAliasName & "','" & strCustomer_code & "','" & strCustomerName & "','" & strCustType & "'," & Val(strSale_of_day) & "," & Val(dblMonth_Budget) & "," & Val(dblMonth_Actual) & "," & Val(dblMonth_Actual) - Val(dblMonth_Budget) & "," & Val(dblYTM_Budget) & "," & Val(dblYTM_Actual) & "," & Val(dblYTM_Actual) - Val(dblYTM_Budget) & ",'" & strtype & "'," & Val(dblSaleValue) & "," & Val(dblCurrentRatePerDay) & "," & Val(dblAskingRatePerDay) & ",'" & gstrUNITID & "')"
                                objCom.CommandText = STR
                                objCom.ExecuteNonQuery()
                            End If
                        Next
                    End With
                End If
                '------------------------------------------------------------------------------------
                objTrans.Commit()
                objCom = Nothing
                objconn.Close()
                objconn = Nothing
                objTrans = Nothing
                MsgBox("Customer Adjustment Saved Successfully With Tracker No.: " & txtTrackerNo.Text, MsgBoxStyle.Information, ResolveResString(100))
                Disable_all()
            End If
        Catch ex As Exception
            objTrans.Rollback()
            objconn.Close()
            objconn = Nothing
            objTrans = Nothing
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function validate_Before_save() As Boolean
        Dim intloop As Int16
        Dim strCustomer_code As String, strtype As String, stramt As String
        Dim dblSaleValue As Object
        Try
            If fpsGrid1.MaxRows > 0 Then
                With Grid3
                    For intloop = 1 To .MaxRows
                        strCustomer_code = Nothing
                        .GetText(enmGrid3.col_Cust_code, intloop, strCustomer_code)
                        strtype = Nothing
                        .GetText(enmGrid3.col_Type, intloop, strtype)
                        stramt = Nothing
                        .GetText(enmGrid3.col_adjust, intloop, stramt)
                        stramt = IIf(IsNothing(stramt), 0, stramt)
                        If strCustomer_code.Length = 0 Then
                            MsgBox("Customer Code Is Missing. Please Select Customer Code For Adjustment.", MsgBoxStyle.Information, ResolveResString(100))
                            validate_Before_save = False
                            Exit Function
                        End If
                        If strtype.Length = 0 Then
                            MsgBox("Adjustment Type Is Missing. Please Select Adjustment Type.", MsgBoxStyle.Information, ResolveResString(100))
                            validate_Before_save = False
                            Exit Function
                        End If
                    Next
                End With
                validate_Before_save = True
            Else
                MsgBox("No Data To Save", MsgBoxStyle.Information, ResolveResString(100))
                validate_Before_save = False
            End If
            With ssprCategoryWiseAdjustment
                For intloop = 1 To .MaxRows
                    strCustomer_code = Nothing
                    .GetText(enumsspr.Customer_Code, intloop, strCustomer_code)
                    .Row = intloop
                    .Col = enumsspr.Sale_Value
                    dblSaleValue = Val(.Text.Trim.ToString())
                    .Row = intloop
                    .Col = enumsspr.Adjustment_Type
                    If strCustomer_code.ToString().Trim.Length > 0 And Val(dblSaleValue.ToString()) <> 0 And .Text.Trim.ToString().Length = 0 Then
                        MsgBox("Adjustment Type Is Missing. Please Select Adjustment Type.", MsgBoxStyle.Information, ResolveResString(100))
                        validate_Before_save = False
                        Exit Function
                    End If
                Next
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
            validate_Before_save = False
        End Try
    End Function
    Private Sub Grid3_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles Grid3.ButtonClicked
        Dim strsql() As String
        Dim strCustomer_list As String
        Dim strquery As String
        Try
            With Grid3
                If e.col = enmGrid3.col_help Then
                    strCustomer_list = get_Customer_List()
                    strquery = "SELECT A.CUSTOMER_CODE, CUST_NAME FROM SALES_TRACKING_ALIAS_MAPPING A INNER JOIN CUSTOMER_MST B ON A.CUSTOMER_CODE = B.CUSTOMER_CODE AND A.UNIT_CODE=B.UNIT_CODE where GETDATE() BETWEEN EFF_FROM_DATE AND EFF_TO_DATE AND A.CUSTOMER_CODE NOT IN (" & strCustomer_list & ") AND A.UNIT_CODE='" & gstrUNITID & "' and ((isnull(B.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= B.deactive_date))"
                    strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strquery)
                    If Not (UBound(strsql) <= 0) Then
                        If Not (UBound(strsql) = 0) Then
                            If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                                MsgBox("Customer Code Does Not Exit.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            Else
                                Call .SetText(enmGrid3.col_Cust_code, e.row, strsql(0))
                                Call .SetText(enmGrid3.col_Cust_name, e.row, strsql(1))
                                Call .SetText(enmGrid3.col_adjust, e.row, "")
                                Call .SetText(enmGrid3.col_Total_sale, e.row, "")
0:                              Call .SetText(enmGrid3.col_Total_sale, e.row, Convert.ToString(get_Month_Sale(strsql(0), Convert.ToDateTime(txtDate.Text))))
                                Call .SetText(enmGrid3.col_Year_Sale, e.row, Convert.ToString(get_Year_Sale(strsql(0), Convert.ToDateTime(txtDate.Text))))
                                Call .SetText(enmGrid3.col_Original_Sales_of_day, e.row, Convert.ToString(get_Sale_Of_Day(strsql(0), Convert.ToDateTime(txtDate.Text))))
                                Call .SetText(enmGrid3.col_Sales_of_day, e.row, Convert.ToString(get_Sale_Of_Day(strsql(0), Convert.ToDateTime(txtDate.Text))))
                                lblSUM1.Text = getSUM_Sale_For_Day().ToString
                            End If
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function get_Month_Sale(ByVal strCust_code As String, ByVal dtDate As DateTime) As Double
        Dim Objconn As SqlConnection
        Try
            Objconn = SqlConnectionclass.GetConnection()
            get_Month_Sale = SqlConnectionclass.ExecuteScalar(Objconn, CommandType.Text, "SELECT DBO.UFN_GET_MONTH_SALES('" & gstrUNITID & "','" & strCust_code & "','" & VB6.Format(dtDate, "DD MMM YYYY") & "')")
            Objconn.Close()
            Objconn = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
            Objconn.Close()
            Objconn = Nothing
        End Try
    End Function
    Private Function get_Sale_Of_Day(ByVal strCust_code As String, ByVal dtDate As DateTime) As Double
        Dim Objconn As SqlConnection
        Try
            Objconn = SqlConnectionclass.GetConnection()
            get_Sale_Of_Day = SqlConnectionclass.ExecuteScalar(Objconn, CommandType.Text, "SELECT DBO.UFN_GET_SALES_OF_DAY('" & gstrUNITID & "','" & strCust_code & "','" & VB6.Format(dtDate, "DD MMM YYYY") & "')")
            Objconn.Close()
            Objconn = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
            Objconn.Close()
            Objconn = Nothing
        End Try
    End Function
    Private Function get_Year_Sale(ByVal strCust_code As String, ByVal dtDate As DateTime) As Double
        Dim Objconn As SqlConnection
        Try
            Objconn = SqlConnectionclass.GetConnection()
            get_Year_Sale = SqlConnectionclass.ExecuteScalar(Objconn, CommandType.Text, "SELECT DBO.UFN_GET_YEAR_SALES('" & gstrUNITID & "','" & strCust_code & "','" & VB6.Format(dtDate.Date, "DD MMM YYYY") & "')")
            Objconn.Close()
            Objconn = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
            Objconn.Close()
            Objconn = Nothing
        End Try
    End Function
    Private Sub AddBlankRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddBlankRow.Click
        If fpsGrid1.MaxRows > 0 Then
            Add_Grid3_blankRow()
        Else
            MsgBox("No Unit Wise Sales Data", MsgBoxStyle.Information, ResolveResString(100))
        End If
    End Sub
    Private Sub cmdTrackerHlp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTrackerHlp.Click
        Dim strsql() As String
        Try
            strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT DISTINCT TRACKER_NO," & DateColumnNameInShowList("RUN_DATE") & "  AS DATE FROM CUST_WISE_SALES_BUDGET  WHERE ACTIVE_FLAG = 1 AND FREEZED = 0 AND TRACKER_NO = (SELECT MAX(TRACKER_NO) FROM CUST_WISE_SALES_BUDGET WHERE ACTIVE_FLAG =1 AND UNIT_CODE='" & gstrUNITID & "') AND UNIT_CODE='" & gstrUNITID & "'")
            If Not (UBound(strsql) <= 0) Then
                If Not (UBound(strsql) = 0) Then
                    If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                        MsgBox("No Un-Freezed Sales Budget To Display. To View Last Freezed Data Click On Show Last Freezed Data.", MsgBoxStyle.Information, ResolveResString(100))
                        populateDate_Data("Freezed Date")
                        dtPicker.Enabled = True
                        CmdDelete.Enabled = False
                        Exit Sub
                    Else
                        fpsGrid1.MaxRows = 0
                        Grid3.MaxRows = 0
                        ssprCategoryWiseAdjustment.MaxRows = 0
                        txtTrackerNo.Text = strsql(0)
                        txtDate.Text = strsql(1)
                        txtTrackerNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtDate.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        populate_Grid1(txtTrackerNo.Text.Trim)
                        populate_grid3(txtTrackerNo.Text.Trim)
                        PopulateCategoryWiseAdjustment(txtTrackerNo.Text.Trim)
                        addRowAtEnterKeyPress(1)
                        get_Working_Days()
                        populateDate_Data("Freezed Date")
                        AddBlankRow.Enabled = True
                        cmdSave.Enabled = True
                        cmdFreeze.Enabled = True
                        CmdDelete.Enabled = True
                        CmbItemCategory.Enabled = True
                        CmbItemCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmdShowData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowData.Click
        Try
            If DataExist("SELECT TRACKER_NO FROM VW_FREEZED_TRACKER_EXIST WHERE UNIT_CODE='" & gstrUNITID & "'") Then
                MsgBox("System Will Dispaly Last Freezed Sales Budget.", MsgBoxStyle.Information, ResolveResString(100))
                getLast_freezed_Budget()
                get_Working_Days()
                populateDate_Data("Freezed Date")
                cmdFreeze.Enabled = False
            Else
                MsgBox("No Freezed Sales Budget Exists. Generate New Sales Budget.", MsgBoxStyle.Information, ResolveResString(100))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Disable_all()
        lblTotalSalesValue.Text = "0"
        CmbItemCategory.SelectedIndex = -1
        fpsGrid1.MaxRows = 0
        Grid3.MaxRows = 0
        ssprCategoryWiseAdjustment.MaxRows = 0
        txtTrackerNo.Text = ""
        lblWork_Day_Value.Text = 0
        lblYTD_Work_Day.Text = 0
        txtDate.Text = ""
        txtReportDate.Text = ""
        txtReportTracker.Text = ""
        lblSUM1.Text = "0"
        lblSUM2.Text = "0"
        txtSearch.Text = ""
        CmdDelete.Enabled = False
        cmdSave.Enabled = False
        cmdFreeze.Enabled = False
        CmbItemCategory.Enabled = False
        CmbItemCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
    End Sub
    Private Function get_Customer_List() As String
        Dim intloop As Int16
        Dim strCust_code As String
        Dim strlist As String = ""
        With Grid3
            For intloop = 1 To .MaxRows
                strCust_code = Nothing
                .GetText(enmGrid3.col_Cust_code, intloop, strCust_code)
                If strCust_code.Length > 0 Then
                    strlist = strlist & "'" & strCust_code & "',"
                End If
            Next
        End With
        If strlist.Length > 0 Then
            strlist = Mid(strlist, 1, (strlist.Length - 1))
        Else
            strlist = "'zzzzzzz'"
        End If
        get_Customer_List = strlist
    End Function
    Private Sub Grid3_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles Grid3.Change
        Dim v_data As String, v_operation As String, strcust_code As String
        Dim v_sales_of_Day As String
        Dim v_hidden As String = String.Empty
        Try
            If e.col = enmGrid3.col_Type Then
                v_data = Nothing
                Grid3.GetText(enmGrid3.col_adjust, e.row, v_data)
                v_hidden = Nothing
                Grid3.GetText(enmGrid3.col_adj_hidden, e.row, v_hidden)
                v_operation = Nothing
                Grid3.GetText(enmGrid3.col_Type, e.row, v_operation)
                If Convert.ToDouble(v_hidden) > 0 Then
                    If Mid(v_operation, 5, 1) = "+" Then
                        Adjust_Unit_Wise_Sales((Convert.ToDecimal(v_hidden) / 1000000), "-", strcust_code)
                    End If
                    If Mid(v_operation, 5, 1) = "-" Then
                        Adjust_Unit_Wise_Sales((Convert.ToDecimal(v_hidden) / 1000000), "+", strcust_code)
                    End If
                End If
                Grid3.SetText(enmGrid3.col_adj_hidden, e.row, 0)
                Grid3.SetText(enmGrid3.col_adjust, e.row, 0)
            ElseIf e.col = enmGrid3.col_Cust_code Then
                Grid3.SetText(enmGrid3.col_adjust, e.row, 0)
            ElseIf e.col = enmGrid3.col_adjust Then
                v_data = Nothing
                Grid3.GetText(e.col, e.row, v_data)
                v_operation = Nothing
                Grid3.GetText(enmGrid3.col_Type, e.row, v_operation)
                strcust_code = Nothing
                Grid3.GetText(enmGrid3.col_Cust_code, e.row, strcust_code)
                v_sales_of_Day = Nothing
                Grid3.GetText(enmGrid3.col_Sales_of_day, e.row, v_sales_of_Day)
                v_hidden = Nothing
                Grid3.GetText(enmGrid3.col_adj_hidden, e.row, v_hidden)
                If Convert.ToDouble(v_hidden) > 0 Then
                    If Mid(v_operation, 5, 1) = "+" Then
                        Adjust_Unit_Wise_Sales((Convert.ToDecimal(v_hidden) / 1000000), "-", strcust_code)
                    End If
                    If Mid(v_operation, 5, 1) = "-" Then
                        Adjust_Unit_Wise_Sales((Convert.ToDecimal(v_hidden) / 1000000), "+", strcust_code)
                    End If
                End If
                Grid3.SetText(enmGrid3.col_adj_hidden, e.row, v_data)
                Adjust_Unit_Wise_Sales((Convert.ToDecimal(v_data) / 1000000), Mid(v_operation, 5, 1), strcust_code)
                lblSUM2.Text = getSUM_Adjustment().ToString
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Adjust_Unit_Wise_Sales(ByVal V_amt As Decimal, ByVal V_operation As String, ByVal v_cust_code As String)
        Dim str_Month_Actual As String, str_Year_Actual As String
        Dim str_Month_Budget As String, str_YTD_Budget As String
        Dim dec_Ask_Rate_per_day As Double
        Dim str_Year_Budget As String
        Dim dec_Data As Decimal, V_DATA As Double
        Dim objconn As SqlConnection
        Dim objCommand As New SqlCommand()
        Dim strSale_of_day As String
        Try
            With fpsGrid1
                strSale_of_day = Nothing
                .GetText(enmGrid1.col_Sale_for_Day, 1, strSale_of_day)
                str_Month_Actual = Nothing
                .GetText(enmGrid1.col_MTD_actual, 1, str_Month_Actual)
                str_Year_Actual = Nothing
                .GetText(enmGrid1.col_YTD_Actual, 1, str_Year_Actual)
                str_Month_Budget = Nothing
                .GetText(enmGrid1.col_MTD_Budget, 1, str_Month_Budget)
                str_YTD_Budget = Nothing
                .GetText(enmGrid1.col_YTD_Budget, 1, str_YTD_Budget)
                Select Case V_operation
                    Case "+"
                        dec_Data = Convert.ToDecimal(strSale_of_day) + V_amt
                        .SetText(enmGrid1.col_Sale_for_Day, 1, dec_Data.ToString("0.0000"))
                        dec_Data = Convert.ToDecimal(str_Month_Actual) + V_amt
                        .SetText(enmGrid1.col_MTD_actual, 1, dec_Data.ToString("0.0000"))
                        .SetText(enmGrid1.col_MVariance, 1, (dec_Data - Convert.ToDecimal(str_Month_Budget)).ToString("0.0000"))
                        dec_Data = Convert.ToDecimal(str_Year_Actual) + V_amt
                        .SetText(enmGrid1.col_YTD_Actual, 1, dec_Data.ToString("0.0000"))
                        .SetText(enmGrid1.col_YVariance, 1, (dec_Data - Convert.ToDecimal(str_YTD_Budget)).ToString("0.0000"))
                    Case "-"
                        dec_Data = Convert.ToDecimal(strSale_of_day) - V_amt
                        .SetText(enmGrid1.col_Sale_for_Day, 1, dec_Data.ToString("0.0000"))
                        dec_Data = Convert.ToDecimal(str_Month_Actual) - V_amt
                        .SetText(enmGrid1.col_MTD_actual, 1, dec_Data.ToString("0.0000"))
                        .SetText(enmGrid1.col_MVariance, 1, (dec_Data - Convert.ToDecimal(str_Month_Budget)).ToString("0.0000"))
                        dec_Data = Convert.ToDecimal(str_Year_Actual) - V_amt
                        .SetText(enmGrid1.col_YTD_Actual, 1, dec_Data.ToString("0.0000"))
                        .SetText(enmGrid1.col_YVariance, 1, (dec_Data - Convert.ToDecimal(str_YTD_Budget)).ToString("0.0000"))
                End Select
                V_DATA = Convert.ToDouble(dec_Data / Convert.ToDouble(lblYTD_Work_Day.Text))
                .SetText(enmGrid1.col_current_rate_day, 1, V_DATA.ToString("0.0000"))
                objconn = SqlConnectionclass.GetConnection()
                objCommand.Connection = objconn
                objCommand.CommandText = "SELECT YEAR_BUDGET FROM CUST_WISE_SALES_BUDGET WHERE TRACKER_NO = " & txtTrackerNo.Text & " AND CUSTOMER_CODE = '" & v_cust_code & "' AND ACTIVE_FLAG = 1 AND FREEZED = 0 AND UNIT_CODE='" & gstrUNITID & "'"
                objCommand.CommandType = CommandType.Text
                str_Year_Budget = objCommand.ExecuteScalar
                objCommand = Nothing
                objconn.Close()
                objconn = Nothing
                dec_Ask_Rate_per_day = (Convert.ToDouble(str_Year_Budget) - dec_Data) / (Convert.ToDouble(lblWork_Day_Value.Text) - Convert.ToDouble(lblYTD_Work_Day.Text))
                .SetText(enmGrid1.col_Ask_rate_day, 1, dec_Ask_Rate_per_day.ToString("0.0000"))
                .Col = enmGrid1.col_Sale_for_Day
                .Col2 = enmGrid1.col_Sale_for_Day
                .Row = 1
                .Row2 = 1
                .BlockMode = True
                .BackColor = Color.Khaki
                .BlockMode = False
                .Col = enmGrid1.col_MTD_actual
                .Col2 = enmGrid1.col_MVariance
                .Row = 1
                .Row2 = 1
                .BlockMode = True
                .BackColor = Color.Khaki
                .BlockMode = False
                .Col = enmGrid1.col_YTD_Actual
                .Col2 = enmGrid1.col_YVariance
                .Row = 1
                .Row2 = 1
                .BlockMode = True
                .BackColor = Color.Khaki
                .BlockMode = False
                .Col = enmGrid1.col_current_rate_day
                .Col2 = enmGrid1.col_Ask_rate_day
                .Row = 1
                .Row2 = 1
                .BlockMode = True
                .BackColor = Color.Khaki
                .BlockMode = False
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub get_Working_Days()
        Dim objconn As SqlConnection
        Dim objDR As SqlDataReader
        Dim objCommand As New SqlCommand()
        Dim intTrackerNo As Integer
        Try
            If txtTrackerNo.Text.Trim = "" Then
                intTrackerNo = 0
            Else
                intTrackerNo = Convert.ToInt32(txtTrackerNo.Text)
            End If
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = "SELECT YEAR_WORK_DAYS, YTD_WORK_DAYS FROM SALES_TRACKING_WORKING_DAYS WHERE UNIT_CODE='" & gstrUNITID & "' AND TRACKER_NO = " & intTrackerNo
            objCommand.CommandType = CommandType.Text
            objDR = objCommand.ExecuteReader()
            While objDR.Read
                lblWork_Day_Value.Text = objDR.GetValue(0).ToString
                lblYTD_Work_Day.Text = objDR.GetValue(1).ToString
            End While
            objDR.Close()
            objDR = Nothing
            objCommand = Nothing
            objconn.Close()
            objconn = Nothing
        Catch ex As Exception
            objDR = Nothing
            objCommand = Nothing
            objconn.Close()
            objconn = Nothing
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function UPDATE_CUST_WISE_SALES_BUDGET(ByVal v_cust_code As String, _
                                                 ByVal v_adj_type As String, _
                                                 ByVal v_monthSale As Double, _
                                                 ByVal v_yearSale As Double, _
                                                 ByVal v_SaleOfDay As Double, _
                                                 ByVal v_adj_amt As Double, _
                                                 ByRef v_Conn As SqlConnection, _
                                                 ByRef v_tran As SqlTransaction) As Boolean
        Dim objCommand1 As New SqlCommand()
        Try
            With objCommand1
                .Connection = v_Conn
                .Transaction = v_tran
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_UPDATE_CUST_WISE_SALES_BUDGET"
                .Parameters.Add("@unitcode", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 16).Value = v_cust_code
                .Parameters.Add("@TRACKER_NO", SqlDbType.BigInt).Value = Convert.ToInt64(txtTrackerNo.Text)
                .Parameters.Add("@ADJUST_TYPE", SqlDbType.Char, 1).Value = v_adj_type
                .Parameters.Add("@ADJ_AMT", SqlDbType.Decimal).Value = v_adj_amt
                .Parameters.Add("@MONTH_SALE", SqlDbType.Money).Value = v_monthSale
                .Parameters.Add("@YEAR_SALE", SqlDbType.Money).Value = v_yearSale
                .Parameters.Add("@SALE_OF_DAY", SqlDbType.Decimal).Value = v_SaleOfDay
                .Parameters.Add("@WORK_DAY", SqlDbType.Money).Value = Convert.ToDouble(lblWork_Day_Value.Text)
                .Parameters.Add("@YTD_DAY", SqlDbType.Money).Value = Convert.ToDouble(lblYTD_Work_Day.Text)
                .ExecuteNonQuery()
            End With
            objCommand1 = Nothing
            UPDATE_CUST_WISE_SALES_BUDGET = True
        Catch ex As Exception
            objCommand1 = Nothing
            MsgBox(ex.Message)
            UPDATE_CUST_WISE_SALES_BUDGET = False
        End Try
    End Function
    Private Sub populate_grid3(ByVal str_tracker_no As String)
        Dim objconn As SqlConnection
        Dim objDR As SqlDataReader
        Dim objCommand As New SqlCommand()
        Try
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = "SELECT * FROM DBO.UFN_POPULATE_ALL_CUST('" & gstrUNITID & "'," & str_tracker_no & ")"
            objCommand.CommandType = CommandType.Text
            objDR = objCommand.ExecuteReader()
            While objDR.Read
                Add_Grid3_blankRow()
                Grid3.SetText(enmGrid3.col_Cust_code, Grid3.MaxRows, objDR.GetValue(0))
                Grid3.SetText(enmGrid3.col_Cust_name, Grid3.MaxRows, objDR.GetValue(1))
                Grid3.SetText(enmGrid3.col_Year_Sale, Grid3.MaxRows, objDR.GetValue(2).ToString)
                Grid3.SetText(enmGrid3.col_Total_sale, Grid3.MaxRows, objDR.GetValue(3).ToString)
                Grid3.SetText(enmGrid3.col_Original_Sales_of_day, Grid3.MaxRows, Convert.ToString(get_Sale_Of_Day(objDR.GetValue(0).ToString, getDateForDB(txtDate.Text))))
                Grid3.SetText(enmGrid3.col_Sales_of_day, Grid3.MaxRows, Convert.ToString(objDR.GetValue(5)))
                Grid3.Col = enmGrid3.col_Type
                Grid3.Row = Grid3.MaxRows
                If objDR.GetString(6).ToUpper = "DR.(+)" Then
                    Grid3.TypeComboBoxCurSel = 0
                Else
                    Grid3.TypeComboBoxCurSel = 1
                End If
                Grid3.SetText(enmGrid3.col_adjust, Grid3.MaxRows, objDR.GetValue(7).ToString)
                Grid3.SetText(enmGrid3.col_adj_hidden, Grid3.MaxRows, objDR.GetValue(7).ToString)
                Grid3.SetText(enmGrid3.col_hidden2, Grid3.MaxRows, objDR.GetValue(7).ToString)
            End While
            objDR.Close()
            objDR = Nothing
            objCommand = Nothing
            objconn.Close()
            objconn = Nothing
            Grid3.Col = enmGrid3.col_Cust_code
            Grid3.Col2 = enmGrid3.col_Sales_of_day
            Grid3.Row = 1
            Grid3.Row2 = Grid3.MaxRows
            Grid3.BlockMode = True
            Grid3.Lock = True
            Grid3.BlockMode = False
            lblSUM1.Text = getSUM_Sale_For_Day().ToString
            lblSUM2.Text = getSUM_Adjustment().ToString
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmdFreeze_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdFreeze.Click
        Dim objconn As SqlConnection
        Dim objTrans As SqlTransaction
        Dim objCom As New SqlCommand()
        Dim str As String
        Try
            If txtTrackerNo.Text.Length = 0 Then
                MsgBox("No Sales Budget To Freeze", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            If IsAdjustment_Exist() Then
                MsgBox("Adjustment To Some Of The Customer Has Not Been Saved." & vbCrLf & _
                         "Please Save The Adjustment By Pressing Save Button And Then Freeze It.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            If MsgBox("Once Freezed, Budget Can Not Be Change." & vbCrLf & "Do You Want To Freeze It?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                objconn = SqlConnectionclass.GetConnection()
                objTrans = objconn.BeginTransaction
                objCom.CommandType = CommandType.Text
                objCom.Connection = objconn
                objCom.Transaction = objTrans
                str = "UPDATE CUST_WISE_SALES_BUDGET SET FREEZED =1, FREEZED_DATE = CONVERT(VARCHAR(12),GETDATE(),106) WHERE UNIT_CODE='" & gstrUNITID & "' and TRACKER_NO =" & txtTrackerNo.Text
                objCom.CommandText = str
                objCom.ExecuteNonQuery()
                str = "IF NOT EXISTS(SELECT TOP 1 1 FROM CUST_WISE_SALES_BUDGET_FREEZED WHERE TRACKER_NO = " & txtTrackerNo.Text & " and UNIT_CODE='" & gstrUNITID & "')" & _
                     " INSERT INTO CUST_WISE_SALES_BUDGET_FREEZED " & _
                     " SELECT * FROM CUST_WISE_SALES_BUDGET WHERE UNIT_CODE='" & gstrUNITID & "' and TRACKER_NO = " & txtTrackerNo.Text
                objCom.CommandText = str
                objCom.ExecuteNonQuery()
                objTrans.Commit()
                MsgBox("Sales Budget Has Been Freezed Successfully.", MsgBoxStyle.Information, ResolveResString(100))
                generateAutoMail()
                objCom = Nothing
                objconn.Close()
                objconn = Nothing
                objTrans = Nothing
                Disable_all()
            End If
        Catch ex As Exception
            objTrans.Rollback()
            objCom = Nothing
            objconn.Close()
            objconn = Nothing
            objTrans = Nothing
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub getLast_freezed_Budget()
        Dim objconn As SqlConnection
        Dim objCom As New SqlCommand()
        Dim strFreezed_tracker As String
        Dim intTrackerNo As Integer
        Try
            objconn = SqlConnectionclass.GetConnection()
            objCom.CommandType = CommandType.Text
            objCom.Connection = objconn
            objCom.CommandText = "SELECT MAX(TRACKER_NO) FROM VW_FREEZED_TRACKER_EXIST where UNIT_CODE='" & gstrUNITID & "'"
            strFreezed_tracker = objCom.ExecuteScalar().ToString()
            If strFreezed_tracker = "" Then
                intTrackerNo = 0
            Else
                intTrackerNo = objCom.ExecuteScalar()
            End If
            txtTrackerNo.Text = strFreezed_tracker
            objCom.CommandText = "select DISTINCT RUN_DATE from CUST_WISE_SALES_BUDGET WHERE UNIT_CODE='" & gstrUNITID & "'  and TRACKER_NO = " & intTrackerNo
            txtDate.Text = VB6.Format(objCom.ExecuteScalar(), gstrDateFormat)
            If strFreezed_tracker.Length > 0 Then
                populate_Grid1(strFreezed_tracker)
                populate_grid3(strFreezed_tracker)
                PopulateCategoryWiseAdjustment(strFreezed_tracker)
                With Grid3
                    .Col = enmGrid3.col_Cust_code
                    .Col2 = enmGrid3.col_adjust
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
                With ssprCategoryWiseAdjustment
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
                CmbItemCategory.Enabled = False
                CmbItemCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdSave.Enabled = False
                cmdFreeze.Enabled = False
                AddBlankRow.Enabled = False
                txtTrackerNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtDate.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                dtPicker.Enabled = True
            End If
            objconn.Close()
            objCom = Nothing
            objconn = Nothing
        Catch ex As Exception
            objconn.Close()
            objCom = Nothing
            objconn = Nothing
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmdGenerateBudget_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerateBudget.Click
        Dim Objconn As SqlConnection
        Dim objCommand As New SqlCommand()
        Try
            If DataExist("SELECT TOP 1 1 FROM CUST_WISE_SALES_BUDGET  WHERE ACTIVE_FLAG = 1 AND FREEZED = 0 AND TRACKER_NO = (SELECT MAX(TRACKER_NO) FROM CUST_WISE_SALES_BUDGET WHERE ACTIVE_FLAG =1 and UNIT_CODE='" & gstrUNITID & "') and UNIT_CODE='" & gstrUNITID & "'") Then
                MsgBox("UnFreezed Sales Data Already Exist. Cannot Generate New Sales Budget", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            fpsGrid1.MaxRows = 0
            Grid3.MaxRows = 0
            ssprCategoryWiseAdjustment.MaxRows = 0
            lblSUM1.Text = "0"
            lblSUM2.Text = "0"
            Objconn = SqlConnectionclass.GetConnection()
            With objCommand
                .Connection = Objconn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_CUST_WISE_SALES_BUDGET"
                .Parameters.Add("@unitcode", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@BUDGETREVISION", SqlDbType.Bit).Value = 1
                .Parameters.Add("@MANUAL_RUN_DATE", SqlDbType.DateTime).Value = getDateForDB(dtPicker.Value)
                Dim p As SqlParameter = .Parameters.Add("@RET_TRACKER_NO", SqlDbType.BigInt, 0)
                p.Value = 0
                p.Direction = ParameterDirection.Output
                .ExecuteNonQuery()
                If .Parameters(.Parameters.Count - 1).Value > 0 Then
                    txtTrackerNo.Text = .Parameters(.Parameters.Count - 1).Value
                    .CommandType = CommandType.Text
                    .CommandText = "select DISTINCT RUN_DATE from CUST_WISE_SALES_BUDGET WHERE UNIT_CODE= '" & gstrUNITID & "' AND TRACKER_NO = " & txtTrackerNo.Text
                    txtDate.Text = VB6.Format(.ExecuteScalar(), gstrDateFormat)
                    populate_Grid1(txtTrackerNo.Text.Trim)
                    populate_grid3(txtTrackerNo.Text.Trim)
                    PopulateCategoryWiseAdjustment(txtTrackerNo.Text.Trim)
                    get_Working_Days()
                    AddBlankRow.Enabled = True
                    cmdFreeze.Enabled = True
                    addRowAtEnterKeyPress(1)
                    CmbItemCategory.Enabled = True
                    CmbItemCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                ElseIf .Parameters(.Parameters.Count - 1).Value = 0 Then
                    MsgBox("Sales Tracker Already Exist For Selected Date: " & dtPicker.Value & vbCrLf & _
                            "Please Select Different Date.", MsgBoxStyle.Information, ResolveResString(100))
                ElseIf .Parameters(.Parameters.Count - 1).Value < 0 Then
                    MsgBox("Month End Sales Tracker Is Not Freezed." & vbCrLf & _
                            "Please Freezed Month End Sales Tracker Before Generating For New Month.", MsgBoxStyle.Information, ResolveResString(100))
                End If
            End With
            objCommand = Nothing
            Objconn.Close()
            Objconn = Nothing
            cmdSave.Enabled = True
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            objCommand = Nothing
            If Objconn.State = ConnectionState.Open Then Objconn.Close()
            Objconn = Nothing
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Dim objCommand1 As New SqlCommand()
        Dim objconn As SqlConnection
        Try
            If txtReportTracker.Text.Length = 0 Then
                MsgBox("Please Select Freezed Sales Budget Tracker No. To Dispaly Report.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            If Not DataExist("SELECT TOP 1 1 FROM CUST_WISE_SALES_BUDGET_FREEZED WHERE UNIT_CODE='" & gstrUNITID & "' AND FREEZED = 1 AND TRACKER_NO = " & txtReportTracker.Text) Then
                MsgBox("Selected Sales Budget Tracker No. Is Not Freezed." & vbCrLf & "Report Can Only Be Displayed For Freezed Tracker.", MsgBoxStyle.Information, ResolveResString(100))
                txtReportDate.Text = ""
                txtReportTracker.Text = ""
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            objconn = SqlConnectionclass.GetConnection()
            With objCommand1
                .Connection = objconn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_CUST_WISE_SALES_BUDGET_REPORT"
                .Parameters.Add("@unitcode", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@TRACKER_NO", SqlDbType.BigInt).Value = Convert.ToInt64(txtReportTracker.Text)
                .Parameters.Add("@IP_ADD", SqlDbType.VarChar, 16).Value = gstrIpaddressWinSck
                .ExecuteNonQuery()
            End With
            objCommand1 = Nothing
            objconn.Close()
            objconn = Nothing
            'With rptSaleTracker
            '    .Reset()
            '    .DiscardSavedData = True
            '    .Connect = gstrREPORTCONNECT
            '    .WindowShowPrintSetupBtn = True
            '    .WindowShowCloseBtn = True
            '    .WindowShowCancelBtn = True
            '    .WindowShowPrintBtn = True
            '    .WindowShowExportBtn = False
            '    .WindowShowSearchBtn = True
            '    .WindowState = Crystal.WindowStateConstants.crptMaximized
            '.WindowTitle = "Sale Budget Tracker Report"
            '.set_Formulas(0, "CompanyName='" & gstrCOMPANY & "'")
            '.ReportFileName = My.Application.Info.DirectoryPath & "\Reports\rptSale_Budget_Tracker.rpt"
            '.SelectionFormula = "{CUST_WISE_SALES_BUDGET_REPORT.TRACKER_NO} = " & txtReportTracker.Text & " AND {CUST_WISE_SALES_BUDGET_REPORT.IP_ADD} ='" & gstrIpaddressWinSck & "'"
            '.Action = 1
            ' End With
            Dim RdAddSold As ReportDocument
            Dim Frm As New eMProCrystalReportViewer
            RdAddSold = Frm.GetReportDocument()
            Frm.ReportHeader = "Sale Budget Tracker Report"
            With RdAddSold
                .Load(My.Application.Info.DirectoryPath & "\Reports\rptSale_Budget_Tracker.rpt")
                .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
                .RecordSelectionFormula = "{CUST_WISE_SALES_BUDGET_REPORT.TRACKER_NO} = " & txtReportTracker.Text & " AND {CUST_WISE_SALES_BUDGET_REPORT.IP_ADD} ='" & gstrIpaddressWinSck & "' and {CUST_WISE_SALES_BUDGET_REPORT.UNIT_CODE} ='" & gstrUNITID & "'"
                Frm.Show()
            End With
        Catch ex As Exception
            objconn = Nothing
            MsgBox(ex.Message)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub generateAutoMail()
        Dim objCommand1 As New SqlCommand()
        Dim objconn As SqlConnection
        Try
            objconn = SqlConnectionclass.GetConnection()
            With objCommand1
                .Connection = objconn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_SALES_BUDGET_AUTOMAIL"
                .Parameters.Add("@unitcode", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@TRACKER_NO", SqlDbType.BigInt).Value = Convert.ToInt64(txtTrackerNo.Text)
                .ExecuteNonQuery()
            End With
            objCommand1 = Nothing
            objconn.Close()
            objconn = Nothing
        Catch ex As Exception
            objCommand1 = Nothing
            objconn.Close()
            objconn = Nothing
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub populateDate_Data(ByVal strDateType As String)

        Dim objconn As SqlConnection
        Dim objCom As New SqlCommand()
        Dim dtFinStartDate As Date

        Try
            Select Case strDateType
                Case "Load"
                    dtPicker.Format = DateTimePickerFormat.Custom
                    dtPicker.CustomFormat = gstrDateFormat
                    dtPicker.Value = GetServerDate()
                    dtPicker.MaxDate = GetServerDate()
                Case "Freezed Date"
                    objconn = SqlConnectionclass.GetConnection()
                    objCom.CommandType = CommandType.Text
                    objCom.Connection = objconn
                    objCom.CommandText = "SELECT MAX(RUN_DATE) AS FREEZED_DATE FROM CUST_WISE_SALES_BUDGET WHERE FREEZED = 1 and UNIT_CODE='" & gstrUNITID & "'"
                    If IsDBNull(objCom.ExecuteScalar()) Then
                        dtPicker.MinDate = SqlConnectionclass.ExecuteScalar("select Fin_Start_Date From Financial_Year_Tb where GETDATE() between Fin_Start_date And Fin_End_date And UNIT_CODE ='" & gstrUNITID & "'")
                    Else
                        dtPicker.MinDate = Convert.ToDateTime(objCom.ExecuteScalar())
                    End If
                    objconn.Close()
                    dtPicker.MaxDate = GetServerDate()
                    dtPicker.Enabled = True
            End Select
            objCom = Nothing
            objconn = Nothing
        Catch ex As Exception
            objconn.Close()
            objCom = Nothing
            objconn = Nothing
        End Try
    End Sub
    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        Disable_all()
        SetSpreadProperty()
    End Sub
    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        Dim objExcel As New Excel.Application
        Dim objwBook As Excel.Workbook
        Dim objwSheet As Excel.Worksheet
        Dim objDT As DataTable
        Dim objDC As DataColumn
        Dim objDR As DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        Dim strFile_Name As String
        Try
            If Len(txtTrackerNo.Text) = 0 Then
                MsgBox("Pls. Select Freezed Tracker No. To Export Data.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            Else
                If Not DataExist("SELECT TOP 1 1 FROM CUST_WISE_SALES_BUDGET_FREEZED WHERE UNIT_CODE='" & gstrUNITID & "' and TRACKER_NO = " & txtTrackerNo.Text) Then
                    MsgBox("Selected Tracker No. Is Not Freezed. Cann't Export Data.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            objDT = SqlConnectionclass.GetDataTable("SELECT UNIT,ALIAS_CODE,ALIAS_NAME,CUSTOMER_CODE,CUSTOMER_NAME,CUST_TYPE,SALES_FOR_DAY,MTD_BUDGET,MTD_ACTUAL,YTD_BUDGET,YTD_ACTUAL,FREEZED_DATE FROM CUST_WISE_SALES_BUDGET_FREEZED WHERE UNIT_CODE='" & gstrUNITID & "' and TRACKER_NO = " & txtTrackerNo.Text)
            'objDT.Columns("FREEZED_DATE").DataType = System.Type.GetType("System.String")
            objwBook = objExcel.Workbooks.Add()
            objwSheet = objwBook.ActiveSheet()
            For Each objDC In objDT.Columns
                colIndex = colIndex + 1
                'NumberFormat = "@"
                objExcel.Cells(1, colIndex) = objDC.ColumnName
            Next
            ProgressBar1.Maximum = objDT.Rows.Count
            For Each objDR In objDT.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                ProgressBar1.Value = ProgressBar1.Value + 1
                For Each objDC In objDT.Columns
                    colIndex = colIndex + 1
                    If objDC.ColumnName = "FREEZED_DATE" Then
                        objExcel.Cells(rowIndex + 1, colIndex) = objDR(objDC.ColumnName)
                        objExcel.Cells(rowIndex + 1, colIndex).numberformat = gstrDateFormat
                    Else
                        objExcel.Cells(rowIndex + 1, colIndex) = objDR(objDC.ColumnName)
                    End If

                Next
            Next
            objwSheet.Columns.AutoFit()
            strFile_Name = gstrLocalCDrive & txtTrackerNo.Text & ".xls"
            Dim blnFileOpen As Boolean = False
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFile_Name)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
            End Try
            If System.IO.File.Exists(strFile_Name) Then
                System.IO.File.Delete(strFile_Name)
            End If
            objwBook.SaveAs(strFile_Name)
            objExcel.Workbooks.Open(strFile_Name)
            objExcel.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
            objDT.Dispose()
        Finally
            ProgressBar1.Value = 0
            ProgressBar1.Maximum = 0
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub cmdReportTracker_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportTracker.Click
        Dim strsql() As String
        Try
            strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT DISTINCT TRACKER_NO," & DateColumnNameInShowList("RUN_DATE") & " as RUN_DATE FROM CUST_WISE_SALES_BUDGET WHERE UNIT_CODE='" & gstrUNITID & "' and FREEZED = 1 ORDER BY RUN_DATE")
            If Not (UBound(strsql) <= 0) Then
                If Not (UBound(strsql) = 0) Then
                    If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                        MsgBox("Freezed Sales Budget Does Not Exist. Generate New Sales Budget.", MsgBoxStyle.Information, ResolveResString(100))
                        populateDate_Data("Freezed Date")
                        dtPicker.Enabled = True
                        Exit Sub
                    Else
                        txtReportTracker.Text = strsql(0)
                        txtReportDate.Text = strsql(1)
                        cmdPrint.Enabled = True
                        cmdPrint.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub txtReportTracker_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtReportTracker.KeyPress
        Try
            If Not (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) And Not (Asc(e.KeyChar) = 8) Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmdAliasReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAliasReport.Click
        Dim objCommand1 As New SqlCommand()
        Dim objconn As SqlConnection
        Try
            If txtReportTracker.Text.Length = 0 Then
                MsgBox("Please Select Freezed Sales Budget Tracker No. To Dispaly Report.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            If Not DataExist("SELECT TOP 1 1 FROM CUST_WISE_SALES_BUDGET_FREEZED WHERE UNIT_CODE='" & gstrUNITID & "' and FREEZED = 1 AND TRACKER_NO = " & txtReportTracker.Text) Then
                MsgBox("Selected Sales Budget Tracker No. Is Not Freezed." & vbCrLf & "Report Can Only Be Displayed For Freezed Tracker.", MsgBoxStyle.Information, ResolveResString(100))
                txtReportDate.Text = ""
                txtReportTracker.Text = ""
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            objconn = SqlConnectionclass.GetConnection()
            With objCommand1
                .Connection = objconn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_CUST_WISE_SALES_BUDGET_REPORT"
                .Parameters.Add("@unitcode", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@TRACKER_NO", SqlDbType.BigInt).Value = Convert.ToInt64(txtReportTracker.Text)
                .Parameters.Add("@IP_ADD", SqlDbType.VarChar, 16).Value = gstrIpaddressWinSck
                .ExecuteNonQuery()
            End With
            objCommand1 = Nothing
            objconn.Close()
            objconn = Nothing
            'With rptSaleTracker
            '    .Reset()
            '    .DiscardSavedData = True
            '    .Connect = gstrREPORTCONNECT
            '    .WindowShowPrintSetupBtn = True
            '    .WindowShowCloseBtn = True
            '    .WindowShowCancelBtn = True
            '    .WindowShowPrintBtn = True
            '    .WindowShowExportBtn = False
            '    .WindowShowSearchBtn = True
            '    .WindowState = Crystal.WindowStateConstants.crptMaximized
            '    .WindowTitle = "Alias Wise Sale Tracker Report"
            '    .set_Formulas(0, "CompanyName='" & gstrCOMPANY & "'")
            '    .ReportFileName = My.Application.Info.DirectoryPath & "\Reports\rptAlias_Wise_Sale_Budget_Tracker.rpt"
            '    .SelectionFormula = "{CUST_WISE_SALES_BUDGET_REPORT.TRACKER_NO} = " & txtReportTracker.Text & " AND {CUST_WISE_SALES_BUDGET_REPORT.IP_ADD} ='" & gstrIpaddressWinSck & "'"
            '    .Action = 1
            'End With
            Dim RdAddSold As ReportDocument
            Dim Frm As New eMProCrystalReportViewer
            RdAddSold = Frm.GetReportDocument()
            Frm.ReportHeader = "Alias Wise Sale Tracker Report"
            With RdAddSold
                .Load(My.Application.Info.DirectoryPath & "\Reports\rptAlias_Wise_Sale_Budget_Tracker.rpt")
                .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
                .RecordSelectionFormula = "{CUST_WISE_SALES_BUDGET_REPORT.TRACKER_NO} = " & txtReportTracker.Text & " AND {CUST_WISE_SALES_BUDGET_REPORT.IP_ADD} ='" & gstrIpaddressWinSck & "' AND {CUST_WISE_SALES_BUDGET_REPORT.UNIT_CODE} ='" & gstrUNITID & "'"
                Frm.Show()
            End With
        Catch ex As Exception
            objconn = Nothing
            MsgBox(ex.Message)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Function getSUM_Sale_For_Day() As Double
        Dim intLoop As Int16
        Dim v_data As String = String.Empty
        Dim dblSum As Double = 0.0
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Grid3
                For intLoop = 1 To .MaxRows
                    v_data = Nothing
                    .GetText(enmGrid3.col_Sales_of_day, intLoop, v_data)
                    dblSum += Convert.ToDouble(v_data)
                Next
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            getSUM_Sale_For_Day = dblSum
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Function getSUM_Adjustment() As Double
        Dim intLoop As Int16
        Dim v_data As String = String.Empty
        Dim dblSum As Double = 0.0
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Grid3
                For intLoop = 1 To .MaxRows
                    v_data = Nothing
                    .GetText(enmGrid3.col_adjust, intLoop, v_data)
                    dblSum += Convert.ToDouble(v_data)
                Next
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            getSUM_Adjustment = dblSum
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub txtSearch_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Try
            If txtSearch.Text = String.Empty Then
                cmdSearch_Click(cmdSearch, New System.EventArgs())
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
        Dim intloop As Int16
        Dim strCust_code As Object = Nothing
        Dim strCust_name As Object = Nothing
        Try
            With Grid3
                For intloop = 1 To .MaxRows
                    strCust_code = Nothing
                    strCust_name = Nothing
                    .GetText(enmGrid3.col_Cust_code, intloop, strCust_code)
                    .GetText(enmGrid3.col_Cust_name, intloop, strCust_name)
                    If (UCase(strCust_code.ToString().Trim) Like UCase(txtSearch.Text.ToString().Trim) & "*") _
                       Or (UCase(strCust_name.ToString().Trim) Like UCase(txtSearch.Text.ToString().Trim) & "*") Then
                        If txtSearch.Text.Length > 0 Then
                            .TopRow = intloop
                            Exit For
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function IsAdjustment_Exist() As Boolean
        Dim intloop As Int16
        Dim strData As String = String.Empty
        Try
            IsAdjustment_Exist = False
            With Grid3
                For intloop = 1 To .MaxRows
                    strData = Nothing
                    .GetText(enmGrid3.col_adjust, intloop, strData)
                    If Convert.ToDecimal(strData) > 0 Then
                        IsAdjustment_Exist = True
                        Exit For
                    End If
                Next
            End With
        Catch ex As Exception
            IsAdjustment_Exist = False
            MsgBox(ex.Message)
        End Try
    End Function
#Region "Code Added By Shabbir On Nov 2010"
    Private Sub CmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDelete.Click
        '---------------------------------------------------------
        'Added by:  Shabbir
        'Added On:  15 Nov 2010
        'Purpose :  To remove the unfreezed tracker
        '---------------------------------------------------------
        Dim objconn As SqlConnection = Nothing
        Dim objCom As New SqlCommand()
        Dim objTrans As SqlTransaction = Nothing
        Try
            objconn = SqlConnectionclass.GetConnection()
            objTrans = objconn.BeginTransaction
            objCom.CommandType = CommandType.Text
            objCom.Connection = objconn
            objCom.Transaction = objTrans
            objCom.CommandText = "DELETE FROM CUST_WISE_SALES_BUDGET WHERE TRACKER_NO = " & txtTrackerNo.Text.Trim & " and UNIT_CODE='" & gstrUNITID & "'"
            objCom.ExecuteNonQuery()
            objCom.CommandText = "DELETE FROM SALES_TRACKING_CUSTOMER_ADJUSTMENT WHERE TRACKER_NO = " & txtTrackerNo.Text.Trim & " and UNIT_CODE='" & gstrUNITID & "'"
            objCom.ExecuteNonQuery()
            objCom.CommandText = "DELETE FROM ITEMCATEGORY_WISE_CUST_SALES_ADJST WHERE TRACKER_NO = " & txtTrackerNo.Text.Trim & " and UNIT_CODE='" & gstrUNITID & "'"
            objCom.ExecuteNonQuery()
            objTrans.Commit()
            MessageBox.Show("Unfreezed sales tracker no  - ( " & txtTrackerNo.Text.Trim & " ) has been deleted successfuly !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmdRefresh.PerformClick()
        Catch ex As Exception
            objTrans.Rollback()
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objCom = Nothing
            objconn = Nothing
        End Try
    End Sub
    Private Function PopulateItemCategory() As String
        'Added By   : Shabbir Hussain
        'Added On   : 19 NOV 2010
        Dim oSqlDr As SqlDataReader
        Dim objCommand As New SqlCommand()
        Dim objconn As SqlConnection = Nothing
        PopulateItemCategory = ""
        Try
            CmbItemCategory.Items.Clear()
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = "SELECT DESCR FROM LISTS WHERE KEY1='TRACKERCATEGORY' and UNIT_CODE='" & gstrUNITID & "'"
            objCommand.CommandType = CommandType.Text
            oSqlDr = objCommand.ExecuteReader()
            If oSqlDr.HasRows = True Then
                While oSqlDr.Read
                    mstrCategory = oSqlDr("DESCR").ToString().Trim
                    CmbItemCategory.Items.Add(oSqlDr("DESCR").ToString().Trim)
                End While
            End If
            PopulateItemCategory = mstrCategory
            If oSqlDr.IsClosed = False Then oSqlDr.Close()
            oSqlDr = Nothing
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objconn = Nothing
            objCommand = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Private Sub PopulateCategoryWiseAdjustment(ByVal str_Tracker_No As String)
        'Added By   : Shabbir Hussain
        'Added On   : 19 NOV 2010
        Dim objconn As SqlConnection = Nothing
        Dim objDR As SqlDataReader
        Dim objCommand As New SqlCommand()
        Dim strQuery As String = ""
        Dim dtTmp As DataTable = Nothing
        Dim Dr As DataRow = Nothing




        Try
            tabSalesTracker.SelectedTab = TabItemCategoryWiseAdjustment
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = "SELECT * FROM ITEMCATEGORY_WISE_CUST_SALES_ADJST WHERE TRACKER_NO = " & str_Tracker_No & " and UNIT_CODE='" & gstrUNITID & "' Order By SALES_VALUE DESC"
            objCommand.CommandType = CommandType.Text
            objDR = objCommand.ExecuteReader()
            ssprCategoryWiseAdjustment.MaxRows = 0
            SetSpreadProperty()
            While objDR.Read
                addRowAtEnterKeyPress(1)
                With ssprCategoryWiseAdjustment
                    .Row = .MaxRows
                    .Col = enumsspr.Item_Category
                    .Text = objDR("ITEM_CATEGORY").ToString().Trim
                    .Col = enumsspr.Customer_Code
                    .Text = objDR("CUSTOMER_CODE").ToString().Trim
                    .Col = enumsspr.Customer_Name
                    .Text = objDR("CUSTOMER_CODE").ToString().Trim
                    .Col = enumsspr.Sale_For_Day
                    .Text = Val(objDR("SALES_FOR_DAY"))
                    strQuery = "SELECT * FROM UDF_GETBDGT_ACT_DTL_CAT_WISE('" & gstrUNITID & "','" & objDR("ITEM_CATEGORY").ToString().Trim & "','" & objDR("CUSTOMER_CODE").ToString().Trim & "','" & Convert.ToDateTime(dtPicker.Value).ToString("dd MMM yyyy") & "')"
                    dtTmp = SqlConnectionclass.ExecuteDataset(SqlConnectionclass.GetConnection(), CommandType.Text, strQuery).Tables(0)
                    If dtTmp.Rows.Count > 0 Then
                        For Each Dr In dtTmp.Rows
                            .Row = .MaxRows
                            .Col = enumsspr.Month_Budget
                            .CellTag = Val(Dr("MONTH_BUDGET").ToString().Trim)
                            .Row = .MaxRows
                            .Col = enumsspr.Month_Actual
                            .CellTag = Val(Dr("MONTH_ACTUAL").ToString().Trim)
                            .Row = .MaxRows
                            .Col = enumsspr.YTM_Budget
                            .CellTag = Val(Dr("YTM_BUDGET").ToString().Trim)
                            .Row = .MaxRows
                            .Col = enumsspr.YTM_Actual
                            .CellTag = Val(Dr("YTM_ACTUAL").ToString().Trim)
                        Next
                    Else
                        .Row = .MaxRows
                        .Col = enumsspr.Month_Budget
                        .CellTag = 0
                        .Row = .MaxRows
                        .Col = enumsspr.Month_Actual
                        .CellTag = 0
                        .Row = .MaxRows
                        .Col = enumsspr.YTM_Budget
                        .CellTag = 0
                        .Row = .MaxRows
                        .Col = enumsspr.YTM_Actual
                        .CellTag = 0
                    End If
                    Dr = Nothing
                    If Not dtTmp Is Nothing Then dtTmp.Dispose()
                    .Row = .MaxRows
                    .Col = enumsspr.Month_Budget
                    .Text = Val(objDR("MONTH_BUDGET"))
                    .Row = .MaxRows
                    .Col = enumsspr.Month_Actual
                    .Text = Val(objDR("MONTH_ACTUAL"))
                    .Row = .MaxRows
                    .Col = enumsspr.Month_Variance
                    .Text = Val(objDR("MONTH_VARIANCE"))
                    .Row = .MaxRows
                    .Col = enumsspr.YTM_Budget
                    .Text = Val(objDR("YTM_BUDGET"))
                    .Row = .MaxRows
                    .Col = enumsspr.YTM_Actual
                    .Text = Val(objDR("YTM_ACTUAL"))
                    .Row = .MaxRows
                    .Col = enumsspr.YTM_Variance
                    .Text = Val(objDR("YTM_VARIANCE"))
                    .Row = .MaxRows
                    .Col = enumsspr.Adjustment_Type
                    If objDR("ADJ_TYPE").ToString().Trim.ToUpper = "DR.(+)" Then
                        .TypeComboBoxCurSel = 0
                    Else
                        .TypeComboBoxCurSel = 1
                    End If
                    .Row = .MaxRows
                    .Col = enumsspr.Adjustment_Type
                    .Text = objDR("ADJ_TYPE").ToString().Trim
                    .Row = .MaxRows
                    .Col = enumsspr.Sale_Value
                    .Text = Val(objDR("SALES_VALUE"))
                    .Row = .MaxRows
                    .Col = enumsspr.Current_Rate_Day
                    .Text = Val(objDR("CURR_RATE_PER_DAY"))
                    .Row = .MaxRows
                    .Col = enumsspr.Ask_Rate_Day
                    .Text = Val(objDR("ASKING_RATE_PER_DAY"))
                End With
            End While
            objDR.Close()
            objDR = Nothing
            objCommand = Nothing
            objconn.Close()
            objconn = Nothing
            CalculateSalesValue()
        Catch ex As Exception
            objDR = Nothing
            objCommand = Nothing
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objconn = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub PopulateDefaultCategoryWiseAdjustment(ByVal str_Tracker_No As String)
        'Added By   : Shabbir Hussain
        'Added On   : 19 NOV 2010
        Dim objconn As SqlConnection = Nothing
        Dim objDR As SqlDataReader
        Dim objCommand As New SqlCommand()
        Dim strQuery As String = ""
        Dim dtTmp As DataTable = Nothing
        Dim Dr As DataRow = Nothing
        Try
            strQuery = "SELECT DISTINCT CUSTOMER_CODE FROM ITEM_CATEGORY_WISE_YEARLY_SALES_BUDGET where UNIT_CODE='" & gstrUNITID & "'"
            strQuery = "UNION "
            strQuery = "Select DISTINCT CUSTOMER_CODE FROM ITEMCATEGORY_WISE_CUST_SALES_ADJST where UNIT_CODE='" & gstrUNITID & "'"
            strQuery = "UNION "
            strQuery = "SELECT DISTINCT CUSTOMER_CODE FROM ITEM_CATEGORY_WISE_YEARLY_SALES_ACTUAL where UNIT_CODE='" & gstrUNITID & "'"
            objconn = SqlConnectionclass.GetConnection()
            objCommand.Connection = objconn
            objCommand.CommandText = strQuery
            objCommand.CommandType = CommandType.Text
            objDR = objCommand.ExecuteReader()
            ssprCategoryWiseAdjustment.MaxRows = 0
            SetSpreadProperty()
            While objDR.Read
                addRowAtEnterKeyPress(1)
                With ssprCategoryWiseAdjustment
                    .SetText(enumsspr.Item_Category, .MaxRows, mstrCategory)
                    .SetText(enumsspr.Customer_Code, .MaxRows, objDR("CUSTOMER_CODE").ToString().Trim)
                    .SetText(enumsspr.Customer_Name, .MaxRows, objDR("CUSTOMER_CODE").ToString().Trim)
                    .SetText(enumsspr.Sale_For_Day, .MaxRows, Val(0))
                    strQuery = "SELECT * FROM UDF_GETBDGT_ACT_DTL_CAT_WISE('" & gstrUNITID & "','" & mstrCategory & "','" & objDR("CUSTOMER_CODE").ToString().Trim & "','" & Convert.ToDateTime(dtPicker.Value).ToString("dd MMM yyyy") & "')"
                    dtTmp = SqlConnectionclass.ExecuteDataset(SqlConnectionclass.GetConnection(), CommandType.Text, strQuery).Tables(0)
                    If dtTmp.Rows.Count > 0 Then
                        For Each Dr In dtTmp.Rows
                            Call .SetText(enumsspr.Month_Budget, .MaxRows, Val(Dr("MONTH_BUDGET").ToString().Trim))
                            .Row = .MaxRows
                            .Col = enumsspr.Month_Budget
                            .CellTag = Val(Dr("MONTH_BUDGET").ToString().Trim)
                            Call .SetText(enumsspr.Month_Actual, .MaxRows, Val(Dr("MONTH_ACTUAL").ToString().Trim))
                            .Row = .MaxRows
                            .Col = enumsspr.Month_Actual
                            .CellTag = Val(Dr("MONTH_ACTUAL").ToString().Trim)
                            Call .SetText(enumsspr.Month_Variance, .MaxRows, Val(Dr("MONTH_ACTUAL").ToString().Trim) - Val(Dr("MONTH_BUDGET").ToString().Trim))
                            Call .SetText(enumsspr.YTM_Budget, .MaxRows, Val(Dr("YTM_BUDGET").ToString().Trim))
                            .Row = .MaxRows
                            .Col = enumsspr.YTM_Budget
                            .CellTag = Val(Dr("YTM_BUDGET").ToString().Trim)
                            Call .SetText(enumsspr.YTM_Actual, .MaxRows, Val(Dr("YTM_ACTUAL").ToString().Trim))
                            .Row = .MaxRows
                            .Col = enumsspr.YTM_Actual
                            .CellTag = Val(Dr("YTM_ACTUAL").ToString().Trim)
                            Call .SetText(enumsspr.YTM_Variance, .MaxRows, Val(Dr("YTM_ACTUAL").ToString().Trim) - Val(Dr("YTM_BUDGET").ToString().Trim))
                        Next
                    Else
                        Call .SetText(enumsspr.Month_Budget, .MaxRows, "")
                        .Row = .MaxRows
                        .Col = enumsspr.Month_Budget
                        .CellTag = 0
                        Call .SetText(enumsspr.Month_Actual, .MaxRows, "")
                        .Row = .MaxRows
                        .Col = enumsspr.Month_Actual
                        .CellTag = 0
                        Call .SetText(enumsspr.Month_Variance, .MaxRows, "")
                        Call .SetText(enumsspr.YTM_Budget, .MaxRows, "")
                        .Row = .MaxRows
                        .Col = enumsspr.YTM_Budget
                        .CellTag = 0
                        Call .SetText(enumsspr.YTM_Actual, .MaxRows, "")
                        .Row = .MaxRows
                        .Col = enumsspr.YTM_Actual
                        .CellTag = 0
                        Call .SetText(enumsspr.YTM_Variance, .MaxRows, "")
                    End If
                    Dr = Nothing
                    If Not dtTmp Is Nothing Then dtTmp.Dispose()
                    .Row = .MaxRows
                    .Col = enumsspr.Adjustment_Type
                    .TypeComboBoxCurSel = 0
                    .Text = "Dr.(+)"
                    Call .SetText(enumsspr.Sale_For_Day, .MaxRows, "")
                    Call .SetText(enumsspr.Sale_Value, .MaxRows, "")
                    Call .SetText(enumsspr.Current_Rate_Day, .MaxRows, "")
                    Call .SetText(enumsspr.Ask_Rate_Day, .MaxRows, "")
                End With
            End While
            objDR.Close()
            objDR = Nothing
            objCommand = Nothing
            objconn.Close()
            objconn = Nothing
            CalculateSalesValue()
        Catch ex As Exception
            objDR = Nothing
            objCommand = Nothing
            If objconn.State = ConnectionState.Open Then objconn.Close()
            objconn = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub CalculateSalesValue()
        Dim intRow As Integer
        Try
            lblTotalSalesValue.Text = 0
            With ssprCategoryWiseAdjustment
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = enumsspr.Sale_Value
                    lblTotalSalesValue.Text = Val(lblTotalSalesValue.Text) + Val(.Text)
                Next intRow
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
#End Region
    Private Sub dtPicker_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtPicker.Validating
        Dim objconn As SqlConnection
        Dim objCom As New SqlCommand()
        Try
            objconn = SqlConnectionclass.GetConnection()
            objCom.CommandType = CommandType.Text
            objCom.Connection = objconn
            objCom.CommandText = "SELECT MAX(RUN_DATE) AS FREEZED_DATE FROM CUST_WISE_SALES_BUDGET WHERE FREEZED = 1 AND UNIT_CODE='" & gstrUNITID & "'"
            If IsDBNull(objCom.ExecuteScalar()) Then
                dtPicker.MinDate = SqlConnectionclass.ExecuteScalar("select Fin_Start_Date From Financial_Year_Tb where GETDATE() between Fin_Start_date And Fin_End_date And UNIT_CODE ='" & gstrUNITID & "'")
            Else
                dtPicker.MinDate = Convert.ToDateTime(objCom.ExecuteScalar())
            End If
            objconn.Close()
            dtPicker.MaxDate = GetServerDate()
            dtPicker.Enabled = True
            objCom = Nothing
            objconn = Nothing
        Catch ex As Exception
            objconn.Close()
            objCom = Nothing
            objconn = Nothing
        End Try
    End Sub
    Private Sub ssprCategoryWiseAdjustment_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles ssprCategoryWiseAdjustment.ButtonClicked
        Dim strsql() As String
        Dim strquery As String
        Dim dtTmp As DataTable = Nothing
        Dim Dr As DataRow
        Try
            If CmbItemCategory.Text.ToString().Trim.Length = 0 Then
                MessageBox.Show("First select the item category !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                If CmbItemCategory.Enabled = True Then CmbItemCategory.Focus()
                Exit Sub
            End If
            With ssprCategoryWiseAdjustment
                If e.col = enumsspr.Hlp Then
                    Dim intloop As Int16
                    Dim strCust_code As Object
                    Dim strlist As String = ""
                    For intloop = 1 To .MaxRows
                        strCust_code = Nothing
                        .GetText(enumsspr.Customer_Code, intloop, strCust_code)
                        If intloop <> e.row And strCust_code.ToString().Trim.Length > 0 Then
                            strlist = strlist & "'" & strCust_code & "',"
                        End If
                    Next
                    If strlist.Length > 0 Then
                        strlist = " AND A.CUSTOMER_CODE NOT IN (" & Mid(strlist, 1, (strlist.Length - 1)) & ") "
                    End If
                    'Currently No validation is there for customer help
                    'As per requirment we have to display all the customers
                    strquery = "SELECT A.CUSTOMER_CODE, CUST_NAME FROM SALES_TRACKING_ALIAS_MAPPING A INNER JOIN CUSTOMER_MST B ON A.CUSTOMER_CODE = B.CUSTOMER_CODE AND A.UNIT_CODE=B.UNIT_CODE where A.UNIT_CODE='" & gstrUNITID & "' AND GETDATE() BETWEEN EFF_FROM_DATE AND EFF_TO_DATE  and ((isnull(B.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= B.deactive_date))" & strlist
                    strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strquery)
                    If Not (UBound(strsql) <= 0) Then
                        If Not (UBound(strsql) = 0) Then
                            If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                                MsgBox("Customer Code Does Not Exit.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            Else
                                .Col = enumsspr.Item_Category
                                .Row = e.row
                                .Text = CmbItemCategory.Text.ToString().Trim()
                                .Col = enumsspr.Customer_Code
                                .Row = e.row
                                .Text = strsql(0)
                                .Col = enumsspr.Customer_Name
                                .Row = e.row
                                .Text = strsql(1)
                                'The function in the following calculates the last month freezed budget and actual data for Month and YTM
                                'if exists otherwise it will give opening value for YTM actual and Budget if  uploaded
                                strquery = "SELECT * FROM UDF_GETBDGT_ACT_DTL_CAT_WISE('" & gstrUNITID & "','" & CmbItemCategory.Text.ToString().Trim & "','" & strsql(0).Trim & "','" & Convert.ToDateTime(dtPicker.Value).ToString("dd MMM yyyy") & "')"
                                dtTmp = SqlConnectionclass.ExecuteDataset(SqlConnectionclass.GetConnection(), CommandType.Text, strquery).Tables(0)
                                If dtTmp.Rows.Count > 0 Then
                                    For Each Dr In dtTmp.Rows
                                        .Row = e.row
                                        .Col = enumsspr.Month_Budget
                                        .CellTag = Val(Dr("MONTH_BUDGET").ToString().Trim)
                                        .Text = Val(Dr("MONTH_BUDGET").ToString().Trim)
                                        .Col = enumsspr.Month_Actual
                                        .Row = e.row
                                        .Text = Val(Dr("MONTH_ACTUAL").ToString().Trim)
                                        .CellTag = Val(Dr("MONTH_ACTUAL").ToString().Trim)
                                        .Col = enumsspr.Month_Variance
                                        .Row = e.row
                                        .Text = Val(Dr("MONTH_ACTUAL").ToString().Trim) - Val(Dr("MONTH_BUDGET").ToString().Trim)
                                        .Col = enumsspr.YTM_Budget
                                        .Row = e.row
                                        .Text = Val(Dr("YTM_BUDGET").ToString().Trim)
                                        .CellTag = Val(Dr("YTM_BUDGET").ToString().Trim)
                                        .Col = enumsspr.YTM_Actual
                                        .Row = e.row
                                        .Text = Val(Dr("YTM_ACTUAL").ToString().Trim)
                                        .CellTag = Val(Dr("YTM_ACTUAL").ToString().Trim)
                                        .Col = enumsspr.YTM_Variance
                                        .Row = e.row
                                        .Text = Val(Dr("YTM_ACTUAL").ToString().Trim) - Val(Dr("YTM_BUDGET").ToString().Trim)
                                    Next
                                Else
                                    .Row = e.row
                                    .Col = enumsspr.Month_Budget
                                    .CellTag = 0
                                    .Text = 0
                                    .Row = e.row
                                    .Col = enumsspr.Month_Actual
                                    .CellTag = 0
                                    .Text = 0
                                    .Col = enumsspr.Month_Variance
                                    .Row = e.row
                                    .Text = 0
                                    .Col = enumsspr.YTM_Budget
                                    .Row = e.row
                                    .Text = 0
                                    .CellTag = 0
                                    .Row = e.row
                                    .Col = enumsspr.YTM_Actual
                                    .CellTag = 0
                                    .Text = 0
                                    .Col = enumsspr.YTM_Variance
                                    .Row = e.row
                                    .Text = 0
                                End If
                                Dr = Nothing
                                If Not dtTmp Is Nothing Then dtTmp.Dispose()
                                .Row = e.row
                                .Col = enumsspr.Adjustment_Type
                                .TypeComboBoxCurSel = 0
                                .Text = "Dr.(+)"
                                .Col = enumsspr.Sale_For_Day
                                .Row = e.row
                                .Text = 0
                                .Col = enumsspr.Sale_Value
                                .Row = e.row
                                .Text = 0
                                .Col = enumsspr.Current_Rate_Day
                                .Row = e.row
                                .Text = 0
                                .Col = enumsspr.Ask_Rate_Day
                                .Row = e.row
                                .Text = 0
                                .Row = .MaxRows
                                .Col = enumsspr.Customer_Code
                                If .Text.Trim.Length > 0 Then
                                    addRowAtEnterKeyPress(1)
                                End If
                                .Row = e.row
                                .Col = enumsspr.Sale_Value
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                            End If
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Dr = Nothing
            If Not dtTmp Is Nothing Then dtTmp.Dispose()
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub ssprCategoryWiseAdjustment_ComboSelChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ComboSelChangeEvent) Handles ssprCategoryWiseAdjustment.ComboSelChange
        Try
            If e.row < 1 Then Exit Sub
            With ssprCategoryWiseAdjustment
                .Row = e.row
                .Col = enumsspr.Sale_Value
                .Text = Val(0)
                Call ssprCategoryWiseAdjustment_EditChange(ssprCategoryWiseAdjustment, New AxFPSpreadADO._DSpreadEvents_EditChangeEvent(enumsspr.Sale_Value, e.row))
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub ssprCategoryWiseAdjustment_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles ssprCategoryWiseAdjustment.EditChange
        'Added By   : Shabbir Hussain
        'Added On   : 19 NOV 2010
        Dim dblSalesValue As Double
        Dim dblMonthActual As Double
        Dim dblMonthBudget As Double
        Dim dblYTMActual As Double
        Dim dblYTMBudget As Double
        Dim strType As String
        If e.row < 1 Then Return
        Try
            Select Case e.col
                Case enumsspr.Adjustment_Type
                Case enumsspr.Sale_Value
                    With ssprCategoryWiseAdjustment
                        .Row = e.row
                        .Col = enumsspr.Customer_Code
                        If .Text.ToString().Trim.Length > 0 Then
                            .Row = e.row
                            .Col = enumsspr.Sale_Value
                            dblSalesValue = Val(.Text.Trim)
                            .Row = e.row
                            .Col = enumsspr.Adjustment_Type
                            strType = .Text.Trim.ToUpper
                            .Row = e.row
                            .Col = enumsspr.Sale_For_Day
                            .Text = IIf(strType.Trim = "DR.(+)", dblSalesValue, -1 * dblSalesValue)
                            .Row = e.row
                            .Col = enumsspr.Month_Budget
                            .Text = Val(.CellTag)
                            dblMonthActual = Val(.CellTag)
                            .Row = e.row
                            .Col = enumsspr.Month_Actual
                            .Text = Val(.CellTag) + IIf(strType.Trim = "DR.(+)", dblSalesValue, -1 * dblSalesValue)
                            dblMonthActual = Val(.CellTag) + IIf(strType.Trim = "DR.(+)", dblSalesValue, -1 * dblSalesValue)
                            .Row = e.row
                            .Col = enumsspr.Month_Variance
                            .Text = dblMonthActual - dblMonthBudget
                            .Row = e.row
                            .Col = enumsspr.YTM_Budget
                            .Text = Val(.CellTag)
                            .Row = e.row
                            .Col = enumsspr.YTM_Actual
                            .Text = Val(.CellTag) + dblMonthActual
                            dblYTMActual = Val(.CellTag) + dblMonthActual
                            .Row = e.row
                            .Col = enumsspr.YTM_Variance
                            .Text = dblYTMActual - dblYTMBudget
                        Else
                            .Row = e.row
                            .Col = enumsspr.Sale_Value
                            .Text = 0
                            .Row = e.row
                            .Col = enumsspr.Sale_For_Day
                            .Text = 0
                            .Row = e.row
                            .Col = enumsspr.Month_Budget
                            .Text = 0
                            .CellTag = 0
                            .Row = e.row
                            .Col = enumsspr.Month_Actual
                            .Text = 0
                            .Row = e.row
                            .Col = enumsspr.Month_Variance
                            .Text = 0
                            .Row = e.row
                            .Col = enumsspr.YTM_Budget
                            .Text = 0
                            .CellTag = 0
                            .Row = e.row
                            .Col = enumsspr.YTM_Actual
                            .Text = 0
                            .CellTag = 0
                            .Row = e.row
                            .Col = enumsspr.YTM_Variance
                            .Text = 0
                            .Col = enumsspr.Hlp
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            If .Enabled = True Then .Focus()
                        End If
                    End With
            End Select
            CalculateSalesValue()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
End Class