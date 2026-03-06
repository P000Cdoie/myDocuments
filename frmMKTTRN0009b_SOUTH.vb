Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0009b_SOUTH
    Inherits System.Windows.Forms.Form
    ' -----------------------------------------------------------------------
    ' Copyright (c)     : MIND Ltd.
    ' Form Name         : FrmMKTTRN0009b
    ' Function Name     : LIST OF DOCUMENT WITH BATCH DETAIL TO KNOCK OFF THE INVOICE QTY
    ' DOCUMENT WISE
    ' OR BATCH WISE (USED IN INVOICE ENTRY FORM.)
    ' Created By        : Sandeep Chadha
    ' Created On        : 30-MARCH-2005
    ' Modify Date       : NIL
    ' Revision History  : -
    ' -----------------------------------------------------------------------
    '=====================================================================
    'Revised By     : Amit Kumar (0670)
    'Revision Date  : 30 May 2011
    'Remarks        : Changes Done To Support Multiunit Function
    '=====================================================================

    ''Revised By:       Saurav Kumar
    ''Revised On:       04 Oct 2013
    ''Issue ID  :       10462231 - eMpro ISuite Changes
    '***********************************************************************************************************************************

    Dim mstrRejectionType As String
    Dim mstrVendor_code As String
    Dim mdblTotalQty As Double
    Dim mstrItemCode As String
    Dim mstrItemDescription As String
    Dim mblnBatchWise As Boolean
    Dim mDecimalValue As Short
    Dim mSelectedDoc As String
    Dim mstrCompileDocDetails As String
    Enum InvRej_Detail
        Doc_No = 1
        Batch_No = 2
        Batch_Date = 3
        MaxQuantity = 4
        Quantity = 5
    End Enum

    Sub SetGridColumnsHeaders()

        On Error GoTo ErrHandler ' Error Handler

        With fpRejDetail
            If mblnBatchWise = True Then
                .MaxCols = 5
                .maxRows = 0
                .set_RowHeight(0, 300)
                .Row = 0 : .Col = InvRej_Detail.Doc_No : .Text = "Doc Number" : .set_ColWidth(InvRej_Detail.Doc_No, 1500)
                .Row = 0 : .Col = InvRej_Detail.Batch_No : .Text = "Batch No" : .set_ColWidth(InvRej_Detail.Batch_No, 1500)
                .Row = 0 : .Col = InvRej_Detail.Batch_Date : .Text = "Batch Date" : .set_ColWidth(InvRej_Detail.Batch_Date, 1500)
                .Row = 0 : .Col = InvRej_Detail.MaxQuantity : .Text = "MaxQty" : .ColHidden = True : .set_ColWidth(InvRej_Detail.MaxQuantity, 1500)
                .Row = 0 : .Col = InvRej_Detail.Quantity : .Text = "Quantity" : .set_ColWidth(InvRej_Detail.Quantity, 1500)
            Else
                .MaxCols = 5
                .MaxRows = 0
                .set_RowHeight(0, 300)
                .Row = 0 : .Col = InvRej_Detail.Doc_No : .Text = "Doc Number" : .set_ColWidth(InvRej_Detail.Doc_No, 4000)
                .Row = 0 : .Col = InvRej_Detail.Batch_No : .Text = "N/A" : .set_ColWidth(InvRej_Detail.Batch_No, 0)
                .Row = 0 : .Col = InvRej_Detail.Batch_Date : .Text = "N/A" : .set_ColWidth(InvRej_Detail.Batch_Date, 0)
                .Row = 0 : .Col = InvRej_Detail.MaxQuantity : .Text = "MaxQty" : .ColHidden = True : .set_ColWidth(InvRej_Detail.MaxQuantity, 1500)
                .Row = 0 : .Col = InvRej_Detail.Quantity : .Text = "Quantity" : .set_ColWidth(InvRej_Detail.Quantity, 1500)
            End If

        End With

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Sub SumofQuantity()
        ' Author        : Sandeep Chadha
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Calulate Sum of Quanity
        ' Datetime      : 30-MARCH-2005

        On Error GoTo ErrHandler

        Dim intRow As Short
        Dim dblTotalQty As Double
        Dim varQty As Object

        dblTotalQty = 0
        With fpRejDetail
            For intRow = 1 To .maxRows
                .Row = intRow
                .Col = InvRej_Detail.Quantity
                varQty = Nothing
                .GetText(InvRej_Detail.Quantity, intRow, varQty)
                dblTotalQty = dblTotalQty + Val(varQty)
            Next
        End With
        lblTotalBatchQty.Text = Format_Quantity(dblTotalQty)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Sub AddNewBlankRow()
        On Error GoTo ErrHandler ' Error Handler
        With fpRejDetail
            .MaxRows = .MaxRows + 1
            .set_RowHeight(.MaxRows, 300)
            .Row = .MaxRows : .Col = InvRej_Detail.Doc_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Row = .MaxRows : .Col = InvRej_Detail.Batch_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Row = .MaxRows : .Col = InvRej_Detail.Batch_Date : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Row = .MaxRows : .Col = InvRej_Detail.MaxQuantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Row = .MaxRows : .Col = InvRej_Detail.Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = DecimalValue
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error GoTo ErrHandler ' Error Handler
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        On Error GoTo ErrHandler ' Error Handler
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
        On Error GoTo ErrHandler ' Error Handler
        If CDbl(lblQuantityTobeIssued.Text) <> CDbl(lblTotalBatchQty.Text) Then
            MsgBox("Total Quantity should be equal to " & Format_Quantity(CDbl(lblQuantityTobeIssued.Text)), MsgBoxStyle.Critical, "Empower")
            Exit Sub
        End If
        frmMKTTRN0009_SOUTH.CompileDocDetails = GetDocumentDetail()
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0009b_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler ' Error Handler
        SetBackGroundColorNew(Me, True)
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / 2) - VB6.PixelsToTwipsY(Me.Height) / 2)
        Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 2) - VB6.PixelsToTwipsX(Me.Width) / 2)
        SetGridColumnsHeaders()
        lblItem_code.Text = Item_Code
        lblDesc.Text = Item_Desc
        lblQuantityTobeIssued.Text = Format_Quantity(Val(CStr(TotalQuantityToBeIssues)))
        Call LoadGrid_Details()
        If Len(Trim(CompileDocDetails)) <> 0 Then
            Call ShowBatchDetails(CompileDocDetails)
        End If
        Call SumofQuantity()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ShowBatchDetails(ByRef strCompileString As String) As Object
        '----------------------------------------------------------------------------------
        'Created By     -   Sandeep Chadha
        'Created On     -   08-Feb-2005
        'Arguments      -   StrCompileString
        'Function       -   Populate the Batch Detail
        '----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strallbatches() As String
        Dim strBatch As Object
        Dim strBatchDetail() As String
        Dim strBatDetail As Object
        Dim intRow As Short
        Dim varDoc_No As Object
        Dim varDoc_BatchNo As Object
        If Len(Trim(strCompileString)) <> 0 Then
            strallbatches = Split(strCompileString, "¶")
            For Each strBatch In strallbatches
                If Len(Trim(strBatch)) <> 0 Then
                    strBatchDetail = Split(strBatch, "§")
                    With fpRejDetail
                        For intRow = 1 To .MaxRows
                            varDoc_No = Nothing
                            Call .GetText(InvRej_Detail.Doc_No, intRow, varDoc_No)
                            varDoc_BatchNo = Nothing
                            Call .GetText(InvRej_Detail.Batch_No, intRow, varDoc_BatchNo)
                            If varDoc_No = CObj(strBatchDetail(0)) And varDoc_BatchNo = CObj(strBatchDetail(1)) Then
                                Call .SetText(InvRej_Detail.Quantity, intRow, CObj(strBatchDetail(3)))
                                Exit For
                            End If
                        Next
                    End With
                End If
            Next strBatch
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub frmMKTTRN0009b_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Property RejectionType() As String
        Get
            RejectionType = mstrRejectionType
        End Get
        Set(ByVal Value As String)
            mstrRejectionType = Value
        End Set
    End Property
    Public Property Vendor_code() As String
        Get
            Vendor_code = mstrVendor_code
        End Get
        Set(ByVal Value As String)
            mstrVendor_code = Value
        End Set
    End Property
    Public Property TotalQuantityToBeIssues() As Double
        Get
            TotalQuantityToBeIssues = mdblTotalQty
        End Get
        Set(ByVal Value As Double)
            mdblTotalQty = Value
        End Set
    End Property
    Public Property Item_Code() As String
        Get
            Item_Code = mstrItemCode
        End Get
        Set(ByVal Value As String)
            mstrItemCode = Value
        End Set
    End Property
    Public Property Item_Desc() As String
        Get
            Item_Desc = mstrItemDescription
        End Get
        Set(ByVal Value As String)
            mstrItemDescription = Value
        End Set
    End Property
    Public Property IsTrans_BatchWise() As Boolean
        Get
            IsTrans_BatchWise = mblnBatchWise
        End Get
        Set(ByVal Value As Boolean)
            mblnBatchWise = Value
        End Set
    End Property
    Public Property DecimalValue() As Short
        Get
            DecimalValue = mDecimalValue
        End Get
        Set(ByVal Value As Short)
            mDecimalValue = Value
        End Set
    End Property
    Public Property Selected_DocNo() As String
        Get
            Selected_DocNo = mSelectedDoc
        End Get
        Set(ByVal Value As String)
            mSelectedDoc = Value
        End Set
    End Property
    Public Property CompileDocDetails() As String
        Get
            CompileDocDetails = mstrCompileDocDetails
        End Get
        Set(ByVal Value As String)
            mstrCompileDocDetails = Value
        End Set
    End Property
    Function Format_Quantity(ByRef dblQuanity As Double) As String
        ' Author        : Sandeep Chadha
        ' Arguments     : Quanity
        ' Return Value  : String
        ' Function      : FORMAT THE QUANITY
        ' Datetime      : 31-Mar-2005
        On Error GoTo ErrHandler
        Format_Quantity = Replace(FormatNumber(dblQuanity, DecimalValue), ",", "")
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub LoadGrid_Details()
        ' Author        : Sandeep Chadha
        ' Arguments     : NIL
        ' Return Value  : NIL
        ' Function      : POPULATE THE GRID.
        ' Datetime      : 31-Mar-2005
        On Error GoTo ErrHandler
        Dim rsDoc_Detail As New ClsResultSetDB
        Dim strSql As String
        If IsTrans_BatchWise = True Then
            'IF GRN CONFIG BATCH WISE
            If Len(Trim(Selected_DocNo)) = 0 Then
                'If Customer Ref No are not selected!!!
                If RejectionType = "GRN" Then
                    strSql = "Select A.Doc_No, Rejected_Quantity, B.Batch_No, B.Batch_Date, c.Current_Batch_Qty as Batch_Rejected_qty from GRN_DTL as a  Inner Join ItemBatch_Dtl as b on a.doc_No=b.doc_No and a.Doc_Type=B.Doc_Type and A.Item_code=B.Item_Code AND A.UNIT_CODE=B.UNIT_CODE  Inner Join ItemBatch_MST as C on b.Item_code=c.Item_code and b.batch_No=c.batch_No and b.UNIT_CODE=c.UNIT_CODE and Location_code='01J1'" & " where A.UNIT_CODE='" + gstrUNITID + "' AND a.Item_code='" & Item_Code & "' and c.Current_Batch_Qty > 0"
                Else
                    strSql = "Select a.Doc_No, B.Batch_No, B.Batch_Date, Batch_Qty as Batch_Rejected_Qty from LRN_DTL as a Inner join ItemBatch_DTL as b On A.Doc_Type=B.Doc_Type and A.Doc_No=b.Doc_No  and a.From_Location=b.From_Location and a.UNIT_CODE=b.UNIT_CODE  Inner Join ItemBatch_MST as C on b.Item_code=c.Item_code and b.batch_No=c.batch_No AND b.UNIT_CODE=c.UNIT_CODE and Location_code='01J1'" & " where A.UNIT_CODE='" + gstrUNITID + "' AND a.Item_code='" & Item_Code & "' and Batch_Qty > 0 "
                End If
            Else
                'If Customer Ref No are selected!!!
                If RejectionType = "GRN" Then
                    strSql = "Select A.Doc_No, Rejected_Quantity, B.Batch_No, B.Batch_Date, c.Current_Batch_Qty as Batch_Rejected_qty from GRN_DTL as a  Inner Join ItemBatch_Dtl as b on a.doc_No=b.doc_No and a.Doc_Type=B.Doc_Type and A.Item_code=B.Item_Code and A.UNIT_CODE=B.UNIT_CODE  Inner Join ItemBatch_MST as C on b.Item_code=c.Item_code and b.batch_No=c.batch_No and b.UNIT_CODE=c.UNIT_CODE and Location_code='01J1'" & " where A.UNIT_CODE='" + gstrUNITID + "' AND a.Item_code='" & Item_Code & "' and a.Doc_NO in (" & Selected_DocNo & ") and c.Current_Batch_Qty > 0"
                Else
                    strSql = "SELECT A.DOC_NO, B.BATCH_NO, B.BATCH_DATE,ISNULL(B.BATCH_QTY,0)-ISNULL(G.QUANTITY,0) AS BATCH_REJECTED_QTY " & _
                                     "FROM LRN_DTL AS A Inner Join (SELECT ISNULL(SUM(BATCH_QTY),0) AS BATCH_QTY,BATCH_NO,BATCH_DATE," & _
                                      "Doc_No , doc_type, Item_Code, From_Location,UNIT_CODE FROM ITEMBATCH_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  GROUP BY DOC_TYPE,FROM_LOCATION,Doc_No,Batch_No,Batch_Date,Item_Code,UNIT_CODE)B " & _
                                      "ON a.doc_type = b.doc_type And a.Doc_No = b.Doc_No AND B.ITEM_CODE=A.ITEM_CODE AND A.FROM_LOCATION=B.FROM_LOCATION INNER JOIN " & _
                                      "ITEMBATCH_MST As C ON B.ITEM_CODE=C.ITEM_CODE AND B.BATCH_NO=C.BATCH_NO AND B.UNIT_CODE=C.UNIT_CODE AND LOCATION_CODE='01J1' LEFT OUTER JOIN " & _
                                      "(SELECT REF_DOC_NO,ITEM_CODE,BATCH_NO,ISNULL(SUM(QUANTITY),0) AS QUANTITY FROM MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  CANCEL_FLAG <> 1 AND REJ_TYPE = 2 " & _
                                      "GROUP BY REF_DOC_NO,BATCH_NO,ITEM_CODE Having IsNull(Sum(Quantity), 0) > 0)G ON A.DOC_NO=G.REF_DOC_NO AND B.ITEM_CODE=G.ITEM_CODE " & _
                                      "AND B.BATCH_NO=G.BATCH_NO where a.Item_code='" & Item_Code & "' and a.Doc_No in (" & Selected_DocNo & ") and ISNULL(B.BATCH_QTY,0)-ISNULL(G.QUANTITY,0)>0 AND A.UNIT_CODE='" + gstrUNITID + "'"
                End If
            End If
        Else
            'DOCUMENT WISE!
            If Len(Trim(Selected_DocNo)) = 0 Then
                If RejectionType = "GRN" Then
                    strSql = "Select A.Doc_No, Rejected_Quantity  from GRN_DTL as a WHERE UNIT_CODE='" + gstrUNITID + "' AND a.Item_code='" & Item_Code & "' and Rejected_Quantity  > 0"
                Else
                    strSql = "Select a.Doc_No, Rejected_Quantity from LRN_DTL as a WHERE UNIT_CODE='" + gstrUNITID + "' AND  a.Item_code='" & Item_Code & "' and Rejected_Quantity  > 0 "
                End If
            Else
                If RejectionType = "GRN" Then
                    strSql = "Select A.Doc_No, Rejected_Quantity  from GRN_DTL as a  WHERE UNIT_CODE='" + gstrUNITID + "' AND  a.Item_code='" & Item_Code & "' and a.Doc_NO in (" & Selected_DocNo & ") and Rejected_Quantity  > 0"
                Else
                    strSql = "Select a.Doc_No, Rejected_Quantity from LRN_DTL as a WHERE UNIT_CODE='" + gstrUNITID + "' AND  a.Item_code='" & Item_Code & "' and a.Doc_No in (" & Selected_DocNo & ") and Rejected_Quantity  > 0 "
                End If
            End If
        End If
        rsDoc_Detail.GetResult(strSql)
        If rsDoc_Detail.RowCount > 0 Then
            Do While Not rsDoc_Detail.EOFRecord
                With fpRejDetail
                    AddNewBlankRow()
                    .SetText(InvRej_Detail.Doc_No, .MaxRows, rsDoc_Detail.GetValue("Doc_No"))
                    If IsTrans_BatchWise = True Then
                        .SetText(InvRej_Detail.Batch_No, .MaxRows, rsDoc_Detail.GetValue("Batch_No"))
                        .SetText(InvRej_Detail.Batch_Date, .MaxRows, VB6.Format(rsDoc_Detail.GetValue("Batch_Date"), gstrDateFormat))
                        .SetText(InvRej_Detail.MaxQuantity, .MaxRows, rsDoc_Detail.GetValue("Batch_Rejected_Qty"))
                        .SetText(InvRej_Detail.Quantity, .MaxRows, "0.00")
                    Else
                        .SetText(InvRej_Detail.MaxQuantity, .MaxRows, rsDoc_Detail.GetValue("Rejected_Quantity"))
                        .SetText(InvRej_Detail.Quantity, .MaxRows, "0.00")
                    End If
                    rsDoc_Detail.MoveNext()
                End With
            Loop
        End If
        rsDoc_Detail.ResultSetClose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function GetDocumentDetail() As String
        ' Author        : Sandeep Chadha
        ' Arguments     : NIL
        ' Return Value  : STRING
        ' Function      : CREATE THE COMPILE STRING TO RETURN THE VALUES TO PARENT FORM
        ' Datetime      : 31-Mar-2005
        On Error GoTo ErrHandler
        Dim strCompileString As String
        Dim varDoc_No As Object
        Dim varBatch_No As Object
        Dim varBatch_Date As Object
        Dim varBatchReq As Object
        Dim varQty As Object
        Dim intRow As Short
        With fpRejDetail
            strCompileString = ""
            For intRow = 1 To .MaxRows
                varDoc_No = Nothing
                .GetText(InvRej_Detail.Doc_No, intRow, varDoc_No)
                varBatch_No = Nothing
                .GetText(InvRej_Detail.Batch_No, intRow, varBatch_No)
                varBatch_Date = Nothing
                .GetText(InvRej_Detail.Batch_Date, intRow, varBatch_Date)
                varQty = Nothing
                .GetText(InvRej_Detail.Quantity, intRow, varQty)
                If Val(varQty) > 0 Then
                    strCompileString = strCompileString & varDoc_No & "§" & varBatch_No & "§" & varBatch_Date & "§" & varQty & "¶"
                End If
            Next
        End With
        GetDocumentDetail = strCompileString
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub fpRejDetail_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles fpRejDetail.Change
        ' Author        : Sandeep Chadha
        ' Arguments     : Validate the Quanity
        ' Return Value  : NIL
        ' Datetime      : 31-Mar-2005
        On Error GoTo ErrHandler
        Dim varMaxPossibleQty As Object
        Dim varQtyEnter As Object
        If e.col = InvRej_Detail.Quantity Then
            With fpRejDetail
                varMaxPossibleQty = Nothing
                .GetText(InvRej_Detail.MaxQuantity, e.row, varMaxPossibleQty)
                varQtyEnter = Nothing
                .GetText(InvRej_Detail.Quantity, e.row, varQtyEnter)
                If varMaxPossibleQty < varQtyEnter Then
                    MsgBox("Quantity should not be greater than " & Format_Quantity(Val(varMaxPossibleQty)), MsgBoxStyle.Critical, "Empower")
                    .SetText(InvRej_Detail.Quantity, e.row, "0.00")
                End If
            End With
            'Calculate Sum of Quantity
            SumofQuantity()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
End Class