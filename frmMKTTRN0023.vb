Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0023
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0023.frm(SUB FORM OF FRMMKTTRN0018.frm)
	' Function          :   Used to Display Details of Supplementary invoice
	' Created By        :   Nisha
	' Created On        :   03 Nov, 2003
    ' Revision History  :   Nisha Rai

    'Modified By Nitin Mehta on 16 May 2011
    'Modified to support MultiUnit functionality
	'===================================================================================
	Private Enum enumSuppDetails
		Doc_No = 1
		Rate = 2
		Packing_Amount = 3
		Basic_Amount = 4
		CustMtrl_Amount = 5
		ToolCost_amount = 6
		Accessible_amount = 7
		TotalExciseAmount = 8
		TotalEcessAmount = 9
		Sales_Tax_Amount = 10
		Surcharge_Sales_Tax_Amount = 11
		total_amount = 12
	End Enum
	Dim mCtlHdrItemCode As System.Windows.Forms.ColumnHeader
	Dim mCtlHdrDrawingNo As System.Windows.Forms.ColumnHeader
	Dim mCtlHdrDescription As System.Windows.Forms.ColumnHeader
	Dim intCheckCounter As Short
	Dim mListItemUserId As System.Windows.Forms.ListViewItem
	Dim mstrInvType As String
	Dim mstrInvSubType As String
	Dim mstrItemText As String
	Dim blnExpinv As Boolean
	Dim intIteminSp As Short
	Public Sub AddHeadersOfGrids()
		'*****************************************************************************************
		'Author              - Nisha Rai
		'Create Date         - 03/11/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Set Header Labels in Grid
		'*****************************************************************************************
		On Error GoTo ErrHandler
		With spdInvSuppDetails
            .MaxCols = 12
            .Row = 0
            .Col = enumSuppDetails.Doc_No : .Text = "Invoice No"
            .Row = 0
            .Col = enumSuppDetails.Rate : .Text = "Rate"
            .Row = 0
            .Col = enumSuppDetails.Packing_Amount : .Text = "Packing amount"
            .Row = 0
            .Col = enumSuppDetails.Basic_Amount : .Text = "Basic"
            .Row = 0
            .Col = enumSuppDetails.CustMtrl_Amount : .Text = "Cust Material"
            .Row = 0
            .Col = enumSuppDetails.ToolCost_amount : .Text = "Tool Cost"
            .Row = 0
            .Col = enumSuppDetails.Accessible_amount : .Text = "Accessible Value"
            .Row = 0
            .Col = enumSuppDetails.TotalExciseAmount : .Text = "Excise Amount"
            .Row = 0
            .Col = enumSuppDetails.TotalEcessAmount : .Text = "Ecess Amount"
            .Row = 0
            .Col = enumSuppDetails.Sales_Tax_Amount : .Text = "S Tax"
            .Row = 0
            .Col = enumSuppDetails.Surcharge_Sales_Tax_Amount : .Text = "SS Tax"
            .Row = 0
            .Col = enumSuppDetails.total_amount : .Text = "Total Value"
            .MaxRows = 0
		End With
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
		Exit Sub
	End Sub
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'*****************************************************************************************
		'Author              - Nisha Rai
		'Create Date         - 03/11/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Add Code on Form Unload
		'*****************************************************************************************
		
		Me.Close()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	Private Sub frmMKTTRN0023_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'*****************************************************************************************
		'Author              - Nisha Rai
		'Create Date         - 03/11/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Set its Position Form From Memory
		'*****************************************************************************************
        On Error GoTo ErrHandler

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	Private Sub frmMKTTRN0023_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'*****************************************************************************************
		'Author              - Nisha Rai
		'Create Date         - 03/11/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Unload from from Memory
		'*****************************************************************************************
		On Error GoTo ErrHandler
        'Me = Nothing
        Me.Dispose()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	Public Sub ToDisplayDataInGrid(ByRef pstrCustomerCode As String, ByRef pstrCustomerPartCode As String, ByRef pstrItemCode As String, ByRef pstrRefInvoiceNo As String, ByRef pstrInvoiceNo As String, ByRef pstrLocation_code As Object, ByRef pstrLastSuppInvNo As String)
		'*****************************************************************************************
		'Author              - Nisha Rai
		'Create Date         - 03/11/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Display Data in Grid
		'*****************************************************************************************
		On Error GoTo ErrHandler
		Dim strSQL As String
		Dim rsSuppdetails As ClsResultSetDB
		Dim intLoopCounter As Short
		Dim intmaxLoop As Short
		rsSuppdetails = New ClsResultSetDB
		strSQL = "Select a.Doc_No,a.PrevRate,a.Rate,a.PrevPacking_Amount,a.Packing_Per * Quantity as Packing_Per,"
		strSQL = strSQL & " a.PrevBasic_Amount,a.Basic_Amount,a.PrevAccessible_amount,a.Accessible_amount,"
		strSQL = strSQL & " a.PrevTotalExciseAmount,a.TotalExciseAmount,a.PrevCustMtrl_Amount,a.CustMtrl_Amount,"
		'Code add by Sourabh
		strSQL = strSQL & " Ecess_amount=isnull(a.Ecess_amount,0),PrevEcess_Amount=isnull(a.PrevEcess_Amount,0),"
		'*******************
		strSQL = strSQL & " a.PrevToolCost_amount,a.ToolCost_amount,a.PrevSales_Tax_Amount,a.Sales_Tax_Amount,"
		strSQL = strSQL & " a.PrevSSTAmount,a.SST_Amount,a.Prevtotal_amount,a.Total_amount from SupplementaryInv_hdr b,"
		strSQL = strSQL & " SupplementaryInv_dtl a Where a.refDoc_no = '" & pstrRefInvoiceNo & "'"
        strSQL = strSQL & " and a.Location_code = '" & pstrLocation_code & "' and a.Cust_Item_code = '"
		strSQL = strSQL & pstrCustomerPartCode & "' and a.ITem_code = '" & pstrItemCode & "' and a.Doc_no <> '"
        strSQL = strSQL & pstrInvoiceNo & "' and a.Doc_no = b.Doc_no and a.UNIT_CODE = b.UNIT_CODE and a.Location_code = b.Location_code AND a.UNIT_CODE='" & gstrUNITID & "' order by a.Doc_No"
		rsSuppdetails.GetResult(strSQL)
		If rsSuppdetails.GetNoRows > 0 Then
			With spdInvSuppDetails
				AddHeadersOfGrids()
                .MaxRows = 0
                .MaxRows = 1 : rsSuppdetails.MoveFirst()
                Call .SetText(enumSuppDetails.Doc_No, 1, pstrRefInvoiceNo)
                Call .SetText(enumSuppDetails.Rate, 1, rsSuppdetails.GetValue("PrevRate"))
                Call .SetText(enumSuppDetails.Packing_Amount, 1, rsSuppdetails.GetValue("PrevPacking_amount"))
                Call .SetText(enumSuppDetails.Basic_Amount, 1, rsSuppdetails.GetValue("PrevBasic_Amount"))
                Call .SetText(enumSuppDetails.Accessible_amount, 1, rsSuppdetails.GetValue("PrevAccessible_amount"))
                Call .SetText(enumSuppDetails.TotalExciseAmount, 1, rsSuppdetails.GetValue("PrevTotalExciseAmount"))
                Call .SetText(enumSuppDetails.TotalEcessAmount, 1, rsSuppdetails.GetValue("PrevEcess_amount"))
                Call .SetText(enumSuppDetails.CustMtrl_Amount, 1, rsSuppdetails.GetValue("PrevCustMtrl_Amount"))
                Call .SetText(enumSuppDetails.ToolCost_amount, 1, rsSuppdetails.GetValue("PrevToolCost_amount"))
                Call .SetText(enumSuppDetails.Sales_Tax_Amount, 1, rsSuppdetails.GetValue("PrevSales_Tax_Amount"))
                Call .SetText(enumSuppDetails.Surcharge_Sales_Tax_Amount, 1, rsSuppdetails.GetValue("PrevSSTAmount"))
                .MaxRows = (rsSuppdetails.GetNoRows)
				intmaxLoop = rsSuppdetails.GetNoRows
				rsSuppdetails.MoveFirst()
				For intLoopCounter = 2 To intmaxLoop
                    If Val(pstrLastSuppInvNo) <> Val(rsSuppdetails.GetValue("Doc_no")) Then
                        Call .SetText(enumSuppDetails.Doc_No, intLoopCounter, rsSuppdetails.GetValue("Doc_no"))
                        Call .SetText(enumSuppDetails.Rate, intLoopCounter, rsSuppdetails.GetValue("Rate"))
                        Call .SetText(enumSuppDetails.Packing_Amount, intLoopCounter, rsSuppdetails.GetValue("Packing_Per"))
                        Call .SetText(enumSuppDetails.Basic_Amount, intLoopCounter, rsSuppdetails.GetValue("Basic_Amount"))
                        Call .SetText(enumSuppDetails.Accessible_amount, intLoopCounter, rsSuppdetails.GetValue("Accessible_amount"))
                        Call .SetText(enumSuppDetails.TotalExciseAmount, intLoopCounter, rsSuppdetails.GetValue("TotalExciseAmount"))
                        Call .SetText(enumSuppDetails.TotalEcessAmount, intLoopCounter, rsSuppdetails.GetValue("Ecess_Amount"))
                        Call .SetText(enumSuppDetails.CustMtrl_Amount, intLoopCounter, rsSuppdetails.GetValue("CustMtrl_Amount"))
                        Call .SetText(enumSuppDetails.ToolCost_amount, intLoopCounter, rsSuppdetails.GetValue("ToolCost_amount"))
                        Call .SetText(enumSuppDetails.Sales_Tax_Amount, intLoopCounter, rsSuppdetails.GetValue("Sales_Tax_Amount"))
                        Call .SetText(enumSuppDetails.Surcharge_Sales_Tax_Amount, intLoopCounter, rsSuppdetails.GetValue("SST_Amount"))
                        Call .SetText(enumSuppDetails.total_amount, intLoopCounter, rsSuppdetails.GetValue("total_amount"))
                    End If
					rsSuppdetails.MoveNext()
				Next 
			End With
			Me.ShowDialog()
		Else
			MsgBox("No Supplementary Invoice Available To Dispaly", MsgBoxStyle.Information, "eMPro")
			Exit Sub
		End If
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
End Class