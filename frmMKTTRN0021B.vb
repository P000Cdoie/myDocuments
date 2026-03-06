Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0021B
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0021B.frm
	' Function          :   Used to select items for Export invoice (Multiple SO)
	' Created By        :   Manoj Kr. Vaish
	' Created On        :   28 May 2007
	' Issue ID          :   19992
	'===================================================================================
    '---------------------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    25/05/2011
    'Revision History  -    Changes for Multi Unit
    '---------------------------------------------------------------------------------------
    Public mstrInvType As String
    Public mstrInvSubType As String
    Public mstrCustomerCode As String
    Public mstrLocationCode As String
    Public mdtInvoiceDate As Date
    Public mstrStockLocation As String
    Public mstrMode As String
    Public mstrDocNo As String
    Dim intIteminSp As Short
    Dim mstrItemText As String

    Private Enum GridHeader1
        Mark = 1
        ItemCode = 2
        DrawingNo = 3
        Description = 4
        CustRef = 5
        BalQuantity = 6
        AmendmentNo = 7
        CURRENCY_CODE = 8
        Payment_Term = 9
        Per_Value = 10
        Rate = 11
        Excise = 12
        CurrentStock = 13
        ScheduleQuantity = 14
        Packing = 15
        Cust_Mtrl = 16
        Others = 17
        BinQty = 18
        DecimalAllow = 19
        NOOFDECIMAL = 20
        SaleQuantity = 21
        FromBox = 22
        ToBox = 23
    End Enum
    Const MaxHdrHlpCols As Short = 23
    Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
        On Error GoTo ErrHandler
        Select Case mstrMode
            Case "ADD"
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    ReDim mItemSODtlArr(0)
                    Me.Close()
                End If
            Case "VIEW", "EDIT"
                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOk.Click
        On Error GoTo ErrHandler
        Select Case mstrMode
            Case "ADD", "EDIT"
                Call FillItemSOArray()
                Me.Close()
            Case "VIEW"
                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0021B_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        SetBackGroundColorNew(Me, True)

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 3.6)
        optPartNo.Checked = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddColumnsInSpread()
        On Error GoTo ErrHandler
        With spItems
            .MaxRows = 0
            .MaxCols = MaxHdrHlpCols
            .Row = 0
            .Font = Me.Font
            .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
            .Col = GridHeader1.Mark : .Text = "Mark" : .set_ColWidth(GridHeader1.Mark, 430)
            .Col = GridHeader1.ItemCode : .Text = "Internal Part No." : .set_ColWidth(GridHeader1.ItemCode, 1600)
            .Col = GridHeader1.DrawingNo : .Text = "Cust Part No." : .set_ColWidth(GridHeader1.DrawingNo, 1600)
            .Col = GridHeader1.Description : .Text = "Description" : .set_ColWidth(GridHeader1.Description, 2600)
            .Col = GridHeader1.CustRef : .Text = "Reference No." : .set_ColWidth(GridHeader1.CustRef, 1300)
            .Col = GridHeader1.BalQuantity : .Text = "SO Quantity" : .set_ColWidth(GridHeader1.BalQuantity, 950)
            .Col = GridHeader1.AmendmentNo : .Text = "Amendment No." : .set_ColWidth(GridHeader1.AmendmentNo, 1300)
            .Col = GridHeader1.CURRENCY_CODE : .Text = "Currency Code" : .set_ColWidth(GridHeader1.CURRENCY_CODE, 1000)
            .Col = GridHeader1.Payment_Term : .Text = "Pay Terms" : .set_ColWidth(GridHeader1.Payment_Term, 1000)
            .Col = GridHeader1.Per_Value : .Text = "Per Value" : .set_ColWidth(GridHeader1.Per_Value, 1000)
            .Col = GridHeader1.Rate : .Text = "Rate" : .set_ColWidth(GridHeader1.Rate, 1000)
            .Col = GridHeader1.Excise : .Text = "Excise" : .set_ColWidth(GridHeader1.Excise, 1000)
            .Col = GridHeader1.CurrentStock : .Text = "Current Stock" : .set_ColWidth(GridHeader1.CurrentStock, 1000)
            .Col = GridHeader1.ScheduleQuantity : .Text = "Schedule Quan." : .set_ColWidth(GridHeader1.ScheduleQuantity, 1000)
            .Col = GridHeader1.Packing : .Text = "Packing" : .set_ColWidth(GridHeader1.Packing, 700)
            .Col = GridHeader1.Cust_Mtrl : .Text = "Cust Mtrl" : .set_ColWidth(GridHeader1.Cust_Mtrl, 900)
            .Col = GridHeader1.Others : .Text = "Others" : .set_ColWidth(GridHeader1.Others, 900)
            .Col = GridHeader1.BinQty : .Text = "Bin Qty" : .set_ColWidth(GridHeader1.BinQty, 900)
            .Col = GridHeader1.DecimalAllow : .Text = "Decimal" : .set_ColWidth(GridHeader1.DecimalAllow, 600)
            .Col = GridHeader1.NOOFDECIMAL : .Text = "NoOfDecimal" : .set_ColWidth(GridHeader1.NOOFDECIMAL, 700)
            .Col = GridHeader1.FromBox : .Text = "From Box" : .set_ColWidth(GridHeader1.FromBox, 700)
            .Col = GridHeader1.ToBox : .Text = "To Box" : .set_ColWidth(GridHeader1.ToBox, 700)
            .Col = GridHeader1.SaleQuantity : .Text = "Sale Qty" : .set_ColWidth(GridHeader1.SaleQuantity, 700)
            .Col = GridHeader1.CURRENCY_CODE : .Col2 = GridHeader1.ToBox : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Col = GridHeader1.ItemCode : .Col2 = GridHeader1.ToBox : .BlockMode = True : .Lock = True : .BlockMode = False
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub optDescription_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDescription.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If optDescription.Checked = True Then
                With spItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader1.Description)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If
            txtsearch.Text = ""
            txtsearch.Focus()
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub OptItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItemCode.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If optItemCode.Checked = True Then
                With spItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader1.ItemCode)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If
            txtsearch.Text = ""
            txtsearch.Focus()
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optPartNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPartNo.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If optPartNo.Checked = True Then
                With spItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader1.DrawingNo)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If
            txtsearch.Text = ""
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optReferenceNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optReferenceNo.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If optItemCode.Checked = True Then
                With spItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader1.CustRef)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If
            txtsearch.Text = ""
            txtsearch.Focus()
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Public Function FillSOItemHelp() As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : Boolean
        ' Function      : Fill Multiple SO Item detail
        ' Datetime      : 29 May 2007
        '-------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsSOItem As ADODB.Recordset
        Dim cmd As ADODB.Command
        Dim intRecordCount As Short
        Dim varCustItemCode As Object
        Dim varCustRef As Object
        Dim varAmendmentNo As Object
        Dim intArrCtr As Short
        Dim varItemCode As Object
        FillSOItemHelp = True
        cmd = New ADODB.Command
        rsSOItem = New ADODB.Recordset
        rsSOItem.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        With cmd
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandText = "SO_Item_Hlp"
            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
            .Parameters.Append(.CreateParameter("@CUST_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, mstrCustomerCode))
            .Parameters.Append(.CreateParameter("@INVTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, mstrInvType))
            .Parameters.Append(.CreateParameter("@INVSUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, mstrInvSubType))
            .Parameters.Append(.CreateParameter("@DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, getDateForDB(mdtInvoiceDate)))
            .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4, mstrMode))
            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, mstrDocNo))
            .Parameters.Append(.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
            .let_ActiveConnection(mP_Connection)
            rsSOItem = .Execute
        End With
        If rsSOItem.EOF Then
            MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & mstrStockLocation & "." & vbCrLf & "3. Check Marketing Schedule in Case of Finished\Trading Goods in SO.", MsgBoxStyle.Information, ResolveResString(100))
            FillSOItemHelp = False
            Exit Function
        End If
        If Len(cmd.Parameters(7).Value) > 0 Then
            MsgBox(cmd.Parameters(7).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            cmd = Nothing
            Exit Function
        End If
        Call AddColumnsInSpread()
        Select Case mstrMode
            Case "ADD"
                With spItems
                    Do While Not rsSOItem.EOF
                        Call AddBlankRow()
                        Call .SetText(GridHeader1.Mark, .MaxRows, rsSOItem.Fields("Checked")) : .TypeCheckCenter = True
                        Call .SetText(GridHeader1.ItemCode, .MaxRows, rsSOItem.Fields("Item_Code"))
                        Call .SetText(GridHeader1.DrawingNo, .MaxRows, rsSOItem.Fields("cust_drgNo"))
                        Call .SetText(GridHeader1.Description, .MaxRows, rsSOItem.Fields("CUST_DRG_DESC"))
                        Call .SetText(GridHeader1.CustRef, .MaxRows, rsSOItem.Fields("Cust_Ref"))
                        Call .SetText(GridHeader1.AmendmentNo, .MaxRows, rsSOItem.Fields("Amendment_no"))
                        If Val(rsSOItem.Fields("Balance_Qty").Value) > 0 Then
                            Call .SetText(GridHeader1.BalQuantity, .MaxRows, rsSOItem.Fields("Balance_Qty"))
                        Else
                            Call .SetText(GridHeader1.BalQuantity, .MaxRows, 0)
                        End If
                        Call .SetText(GridHeader1.CURRENCY_CODE, .MaxRows, rsSOItem.Fields("CURRENCY_CODE"))
                        Call .SetText(GridHeader1.Payment_Term, .MaxRows, rsSOItem.Fields("TERM_PAYMENT"))
                        Call .SetText(GridHeader1.Per_Value, .MaxRows, rsSOItem.Fields("PERVALUE"))
                        Call .SetText(GridHeader1.Rate, .MaxRows, rsSOItem.Fields("Rate"))
                        Call .SetText(GridHeader1.Excise, .MaxRows, rsSOItem.Fields("EXCISE_DUTY"))
                        Call .SetText(GridHeader1.CurrentStock, .MaxRows, rsSOItem.Fields("Cur_Bal"))
                        Call .SetText(GridHeader1.ScheduleQuantity, .MaxRows, rsSOItem.Fields("pending_schedule"))
                        Call .SetText(GridHeader1.Packing, .MaxRows, rsSOItem.Fields("Packing_type"))
                        Call .SetText(GridHeader1.Cust_Mtrl, .MaxRows, rsSOItem.Fields("Cust_Mtrl"))
                        Call .SetText(GridHeader1.Others, .MaxRows, rsSOItem.Fields("Others"))
                        Call .SetText(GridHeader1.BinQty, .MaxRows, rsSOItem.Fields("binquantity"))
                        Call .SetText(GridHeader1.DecimalAllow, .MaxRows, IIf(rsSOItem.Fields("DECIMAL_ALLOWED_FLAG").Value, 1, 0))
                        Call .SetText(GridHeader1.NOOFDECIMAL, .MaxCols, rsSOItem.Fields("NOOFDECIMAL"))
                        If Val(rsSOItem.Fields("Balance_Qty").Value) <= 0 Then
                            .Col = -1
                            .BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFFF)
                        Else
                            .BackColor = System.Drawing.Color.White
                        End If
                        rsSOItem.MoveNext()
                    Loop
                End With
                rsSOItem.Close()
                rsSOItem = Nothing
                cmd = Nothing
                Me.ShowDialog()
            Case "VIEW"
                With spItems
                    Do While Not rsSOItem.EOF
                        Call AddBlankRow()
                        Call .SetText(GridHeader1.Mark, .MaxRows, rsSOItem.Fields("Checked")) : .TypeCheckCenter = True
                        Call .SetText(GridHeader1.ItemCode, .MaxRows, rsSOItem.Fields("Item_Code"))
                        Call .SetText(GridHeader1.DrawingNo, .MaxRows, rsSOItem.Fields("CUST_ITEM_CODE"))
                        Call .SetText(GridHeader1.Description, .MaxRows, rsSOItem.Fields("CUST_ITEM_DESC"))
                        Call .SetText(GridHeader1.CustRef, .MaxRows, rsSOItem.Fields("Cust_Ref"))
                        Call .SetText(GridHeader1.AmendmentNo, .MaxRows, rsSOItem.Fields("Amendment_no"))
                        If Val(rsSOItem.Fields("Balance_Qty").Value) > 0 Then
                            Call .SetText(GridHeader1.BalQuantity, .MaxRows, rsSOItem.Fields("Balance_Qty"))
                        Else
                            Call .SetText(GridHeader1.BalQuantity, .MaxRows, 0)
                        End If
                        Call .SetText(GridHeader1.Rate, .MaxRows, rsSOItem.Fields("Rate"))
                        Call .SetText(GridHeader1.Excise, .MaxRows, rsSOItem.Fields("Excise_Type"))
                        Call .SetText(GridHeader1.Packing, .MaxRows, rsSOItem.Fields("Packing_type"))
                        Call .SetText(GridHeader1.Cust_Mtrl, .MaxRows, rsSOItem.Fields("Cust_Mtrl"))
                        Call .SetText(GridHeader1.Others, .MaxRows, rsSOItem.Fields("Others"))
                        Call .SetText(GridHeader1.BinQty, .MaxRows, rsSOItem.Fields("binquantity"))
                        Call .SetText(GridHeader1.SaleQuantity, .MaxRows, rsSOItem.Fields("SALES_QUANTITY"))
                        Call .SetText(GridHeader1.FromBox, .MaxRows, rsSOItem.Fields("From_Box"))
                        Call .SetText(GridHeader1.ToBox, .MaxRows, rsSOItem.Fields("To_Box"))
                        Call .SetText(GridHeader1.CurrentStock, .MaxRows, 0)
                        Call .SetText(GridHeader1.ScheduleQuantity, .MaxRows, 0)
                        .Row = .MaxRows : .Col = GridHeader1.Mark : .Lock = True
                        If Val(rsSOItem.Fields("Balance_Qty").Value) <= 0 Then
                            .Col = -1
                            .BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFFF)
                        Else
                            .BackColor = System.Drawing.Color.White
                        End If
                        rsSOItem.MoveNext()
                    Loop
                End With
                rsSOItem.Close()
                rsSOItem = Nothing
                cmd = Nothing
                Call FillItemSOArray()
            Case "EDIT"
                With spItems
                    Do While Not rsSOItem.EOF
                        Call AddBlankRow()
                        Call .SetText(GridHeader1.Mark, .MaxRows, rsSOItem.Fields("Checked")) : .TypeCheckCenter = True
                        Call .SetText(GridHeader1.ItemCode, .MaxRows, rsSOItem.Fields("Item_Code"))
                        Call .SetText(GridHeader1.DrawingNo, .MaxRows, rsSOItem.Fields("CUST_ITEM_CODE"))
                        Call .SetText(GridHeader1.Description, .MaxRows, rsSOItem.Fields("CUST_ITEM_DESC"))
                        Call .SetText(GridHeader1.CustRef, .MaxRows, rsSOItem.Fields("Cust_Ref"))
                        Call .SetText(GridHeader1.AmendmentNo, .MaxRows, rsSOItem.Fields("Amendment_no"))
                        If Val(rsSOItem.Fields("Balance_Qty").Value) > 0 Then
                            Call .SetText(GridHeader1.BalQuantity, .MaxRows, rsSOItem.Fields("Balance_Qty"))
                        Else
                            Call .SetText(GridHeader1.BalQuantity, .MaxRows, 0)
                        End If
                        Call .SetText(GridHeader1.Payment_Term, .MaxRows, rsSOItem.Fields("TERM_PAYMENT"))
                        Call .SetText(GridHeader1.Rate, .MaxRows, rsSOItem.Fields("Rate"))
                        Call .SetText(GridHeader1.Excise, .MaxRows, rsSOItem.Fields("Excise_Type"))
                        Call .SetText(GridHeader1.Packing, .MaxRows, rsSOItem.Fields("Packing_type"))
                        Call .SetText(GridHeader1.Cust_Mtrl, .MaxRows, rsSOItem.Fields("Cust_Mtrl"))
                        Call .SetText(GridHeader1.Others, .MaxRows, rsSOItem.Fields("Others"))
                        Call .SetText(GridHeader1.BinQty, .MaxRows, rsSOItem.Fields("binquantity"))
                        Call .SetText(GridHeader1.SaleQuantity, .MaxRows, rsSOItem.Fields("SALES_QUANTITY"))
                        Call .SetText(GridHeader1.FromBox, .MaxRows, rsSOItem.Fields("From_Box"))
                        Call .SetText(GridHeader1.ToBox, .MaxRows, rsSOItem.Fields("To_Box"))
                        Call .SetText(GridHeader1.CurrentStock, .MaxRows, rsSOItem.Fields("Cur_Bal"))
                        Call .SetText(GridHeader1.ScheduleQuantity, .MaxRows, rsSOItem.Fields("pending_schedule"))
                        Call .SetText(GridHeader1.DecimalAllow, .MaxRows, IIf(rsSOItem.Fields("DECIMAL_ALLOWED_FLAG").Value, 1, 0))
                        .Row = .MaxRows : .Col = GridHeader1.Mark : .Lock = True
                        If Val(rsSOItem.Fields("Balance_Qty").Value) <= 0 Then
                            .Col = -1
                            .BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFFF)
                        Else
                            .BackColor = System.Drawing.Color.White
                        End If
                        rsSOItem.MoveNext()
                    Loop
                End With
                rsSOItem.Close()
                rsSOItem = Nothing
                cmd = Nothing
                Call FillItemSOArray()
        End Select
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Resume
    End Function
    Private Sub AddBlankRow()
        On Error GoTo ErrHandler
        With spItems
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.MaxRows, 300)
            .Col = GridHeader1.Mark
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.ItemCode
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.DrawingNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Description
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.CustRef
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.AmendmentNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.BalQuantity
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.CURRENCY_CODE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Payment_Term
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Per_Value
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Rate
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Excise
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.CurrentStock
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.ScheduleQuantity
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Packing
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Cust_Mtrl
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.Others
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.BinQty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.DecimalAllow
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.NOOFDECIMAL
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.FromBox
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.ToBox
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = GridHeader1.SaleQuantity
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtsearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        On Error GoTo ErrHandler
        Call SearchItem()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Sub SearchItem()
        On Error GoTo ErrHandler
        Dim intCount As Short
        With spItems
            .Row = -1
            .Col = -1
            .Font = VB6.FontChangeBold(.Font, False)
            If optItemCode.Checked Then
                .Col = 2
            End If
            If optPartNo.Checked Then
                .Col = 3
            End If
            If optDescription.Checked Then
                .Col = 4
            End If
            If optReferenceNo.Checked Then
                .Col = 5
            End If
            If Len(Trim(txtsearch.Text)) > 0 Then
                For intCount = 1 To .MaxRows
                    .Row = intCount
                    If UCase(Mid(.Text, 1, Len(txtsearch.Text))) = UCase(txtsearch.Text) Then
                        .TopRow = .Row
                        .Col = -1
                        .Font = VB6.FontChangeBold(.Font, True)
                        Exit Sub
                    End If
                Next
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSearch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtsearch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtsearch_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtsearch.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        With spItems
            If KeyCode = 13 And Len(Trim(txtsearch.Text)) > 0 Then
                .Col = 1
                .Value = IIf(CBool(.Value), False, True)
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub FillItemSOArray()
        On Error GoTo ErrHandler
        Dim intctr As Short
        Dim intArrIndex As Short
        ReDim mItemSODtlArr(0)
        Select Case mstrMode
            Case "ADD", "EDIT"
                With spItems
                    For intctr = 1 To .MaxRows Step 1
                        .Row = intctr
                        .Col = GridHeader1.Mark
                        If CBool(.Value) Then
                            ReDim Preserve mItemSODtlArr(UBound(mItemSODtlArr) + 1)
                            intArrIndex = UBound(mItemSODtlArr)
                            .Col = GridHeader1.ItemCode
                            mItemSODtlArr(intArrIndex).ItemCode = Trim(.Value)
                            .Col = GridHeader1.DrawingNo
                            mItemSODtlArr(intArrIndex).CustDrgNo = Trim(.Value)
                            .Col = GridHeader1.CustRef
                            mItemSODtlArr(intArrIndex).CustRef = Trim(.Value)
                            .Col = GridHeader1.AmendmentNo
                            mItemSODtlArr(intArrIndex).AmendmentNo = Trim(.Value)
                            .Col = GridHeader1.Rate
                            mItemSODtlArr(intArrIndex).Rate = CDec(.Value)
                            .Col = GridHeader1.Excise
                            mItemSODtlArr(intArrIndex).Excise = Trim(.Value)
                            .Col = GridHeader1.Packing
                            mItemSODtlArr(intArrIndex).Packing = Trim(.Value)
                            .Col = GridHeader1.Cust_Mtrl
                            mItemSODtlArr(intArrIndex).Cust_Mtrl = Val(.Value)
                            .Col = GridHeader1.Others
                            mItemSODtlArr(intArrIndex).Others = Val(.Value)
                            .Col = GridHeader1.BinQty
                            mItemSODtlArr(intArrIndex).BinQty = Val(.Value)
                            .Col = GridHeader1.DecimalAllow
                            mItemSODtlArr(intArrIndex).Decimal_1 = Val(.Value)
                            .Col = GridHeader1.NOOFDECIMAL
                            mItemSODtlArr(intArrIndex).NoOfDeci = Val(.Value)
                            .Col = GridHeader1.CurrentStock
                            mItemSODtlArr(intArrIndex).CurrentStock = Val(.Value)
                            .Col = GridHeader1.ScheduleQuantity
                            mItemSODtlArr(intArrIndex).ScheduleQty = Val(.Value)
                            .Col = GridHeader1.BalQuantity
                            mItemSODtlArr(intArrIndex).balqty = Val(.Value)
                            .Col = GridHeader1.Payment_Term
                            mItemSODtlArr(intArrIndex).CreditTerm = Trim(.Value)
                            .Col = GridHeader1.Per_Value
                            mintPerValue = Val(.Value)
                            .Col = GridHeader1.CURRENCY_CODE
                            mstrCurrency = Trim(.Value)
                            .Col = GridHeader1.Payment_Term
                            mstrCreditTerm = Trim(.Value)
                            If mstrMode = "EDIT" Then
                                .Col = GridHeader1.FromBox
                                mItemSODtlArr(intArrIndex).FromBox = Val(.Value)
                                .Col = GridHeader1.ToBox
                                mItemSODtlArr(intArrIndex).ToBox = Val(.Value)
                                .Col = GridHeader1.SaleQuantity
                                mItemSODtlArr(intArrIndex).SaleQuantity = Val(.Value)
                            End If
                        End If
                    Next intctr
                End With
            Case "VIEW"
                With spItems
                    For intctr = 1 To .MaxRows Step 1
                        .Row = intctr
                        ReDim Preserve mItemSODtlArr(UBound(mItemSODtlArr) + 1)
                        intArrIndex = UBound(mItemSODtlArr)
                        .Col = GridHeader1.ItemCode
                        mItemSODtlArr(intArrIndex).ItemCode = Trim(.Value)
                        .Col = GridHeader1.DrawingNo
                        mItemSODtlArr(intArrIndex).CustDrgNo = Trim(.Value)
                        .Col = GridHeader1.CustRef
                        mItemSODtlArr(intArrIndex).CustRef = Trim(.Value)
                        .Col = GridHeader1.AmendmentNo
                        mItemSODtlArr(intArrIndex).AmendmentNo = Trim(.Value)
                        .Col = GridHeader1.Rate
                        mItemSODtlArr(intArrIndex).Rate = Val(.Value)
                        .Col = GridHeader1.Excise
                        mItemSODtlArr(intArrIndex).Excise = Trim(.Value)
                        .Col = GridHeader1.Packing
                        mItemSODtlArr(intArrIndex).Packing = Trim(.Value)
                        .Col = GridHeader1.Cust_Mtrl
                        mItemSODtlArr(intArrIndex).Cust_Mtrl = Val(.Value)
                        .Col = GridHeader1.Others
                        mItemSODtlArr(intArrIndex).Others = Val(.Value)
                        .Col = GridHeader1.BinQty
                        mItemSODtlArr(intArrIndex).BinQty = Val(.Value)
                        .Col = GridHeader1.FromBox
                        mItemSODtlArr(intArrIndex).FromBox = Val(.Value)
                        .Col = GridHeader1.ToBox
                        mItemSODtlArr(intArrIndex).ToBox = Val(.Value)
                        .Col = GridHeader1.SaleQuantity
                        mItemSODtlArr(intArrIndex).SaleQuantity = Val(.Value)
                        .Col = GridHeader1.BalQuantity
                        mItemSODtlArr(intArrIndex).balqty = Val(.Value)
                    Next intctr
                End With
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spItems_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles spItems.Change
        On Error GoTo ErrHandler
        Call spItems_ClickEvent(spItems, New AxFPSpreadADO._DSpreadEvents_ClickEvent(GridHeader1.Mark, e.row))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub spItems_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles spItems.ClickEvent
        Dim blnFlag As Boolean
        Dim blnflag1 As Boolean
        Dim varExcise As Object
        Dim varCurrencyCode As Object
        Dim varRate As Object
        Dim varPerValue As Object
        Dim varPaymentTrem As Object
        Dim intSubItem As Short
        Dim varDrawingNo As Object
        On Error GoTo ErrHandler
        With spItems
            If e.row > 0 Then
                If e.col = GridHeader1.Mark Then
                    .Row = e.row : .Col = e.col
                    If CBool(.Value) = False Then Exit Sub
                    varExcise = Nothing
                    blnFlag = .GetText(GridHeader1.Excise, e.row, varExcise)
                    varRate = Nothing
                    blnflag1 = .GetText(GridHeader1.Rate, e.row, varRate)
                    varPerValue = Nothing
                    blnflag1 = .GetText(GridHeader1.Per_Value, e.row, varPerValue)
                    varPaymentTrem = Nothing
                    blnflag1 = .GetText(GridHeader1.Payment_Term, e.row, varPaymentTrem)
                    varCurrencyCode = Nothing
                    blnflag1 = .GetText(GridHeader1.CURRENCY_CODE, e.row, varCurrencyCode)
                    varDrawingNo = Nothing
                    blnflag1 = .GetText(GridHeader1.DrawingNo, e.row, varDrawingNo)
                    For intSubItem = 1 To .MaxRows
                        .Row = intSubItem
                        .Col = GridHeader1.Mark
                        '' If Item is Checked
                        If CBool(.Value) = True And .Row <> e.row Then
                            .Col = GridHeader1.CURRENCY_CODE
                            ''If Item is repeated.
                            If UCase(Trim(varCurrencyCode)) = UCase(Trim(.Text)) Then
                                .Col = GridHeader1.Per_Value
                                If Trim(varPerValue) = Trim(.Text) Then
                                    .Col = GridHeader1.Excise
                                    If Trim(varExcise) = Trim(.Text) Then
                                    Else
                                        MsgBox("You can not select items for different Exicse value.", MsgBoxStyle.Information, "eMPro")
                                        .Col = GridHeader1.Mark
                                        .Row = e.row
                                        .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox("You can not select items for different per value.", MsgBoxStyle.Information, "eMPro")
                                    .Col = GridHeader1.Mark
                                    .Row = e.row
                                    .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                                    Exit Sub
                                End If
                            Else
                                MsgBox("You can not select items for different Currency.", MsgBoxStyle.Information, "eMPro")
                                .Col = GridHeader1.Mark
                                .Row = e.row
                                .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class