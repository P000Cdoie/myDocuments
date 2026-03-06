Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0041
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   frmMKTTRN0041.frm
	' Function          :   Used to select Forms for Sales Order
	' Created By        :   Arshad Ali
	' Created On        :   18 March, 2005
    '===================================================================================
    'REVISION HISTORY
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    25/04/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------

	Public mstrFormDetails As String
	Dim mstrParentForm As String
	Dim mstrCust_code As String
    Public INVOICE_NO As String
    Public RefNo As String
    Public InvType As String
    Public RejType As String

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Dim intCount As Short
        Dim blnSelected As Boolean
        Dim varFormType As Object
        Dim varFormNo As Object
        Dim blnCheck As Boolean
        With fpForms
            blnSelected = False
            For intCount = 1 To .MaxRows
                .Row = intCount
                .Col = 1
                If CDbl(.Value) = 1 Then
                    blnSelected = True
                End If
            Next
            If Not blnSelected Then
                If MsgBox("No forms selected." & vbCrLf & "Are you sure to proceed?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "empower") = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If
            InvoiceFormDetails = ""
            For intCount = 1 To .MaxRows
                .Row = intCount
                .Col = 1
                If CBool(.Value) = True Then
                    varFormType = Nothing
                    blnCheck = .GetText(2, intCount, varFormType)
                    varFormNo = Nothing
                    blnCheck = .GetText(3, intCount, varFormNo)
                    InvoiceFormDetails = InvoiceFormDetails & Trim(varFormType) & "|" & Trim(varFormNo) & "^"

                End If
            Next intCount
        End With

        SetFormDetailsTOParentForm()

        Me.Close()

    End Sub

    Sub SetFormDetailsTOParentForm()
        On Error GoTo ErrHandler

        Select Case UCase(ParentForm_Renamed)
            Case ""
                MsgBox("PARENT FORM NAME NOT FOUND CANT SAVE FORM DETALS.", MsgBoxStyle.Critical, "EMPOWER")
            Case "FRMMKTTRN0001"
                frmMKTTRN0001.mstrFormDetails = mstrFormDetails
            Case "FRMMKTTRN0009"
        End Select


        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0041_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Escape Then
            cmdCancel_Click(cmdcancel, New System.EventArgs())
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Public Sub frmMKTTRN0041_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler

        Select Case UCase(ParentForm_Renamed)
            Case "FRMMKTTRN0001"
                fillGrid()
            Case "FRMMKTTRN0009"
                fillGrid_FrmMKTTRN0009()
        End Select

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Public Function Find_Value(ByRef strField As String) As String
        '-------------------------- --------------------------------------------------
        'Author         :   Sandeep Chadha
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

    Public Sub fillGrid_FrmMKTTRN0009()
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim intRow As Short
        Dim rsForm As New ClsResultSetDB
        Dim strPO As String

        If frmMKTTRN0009.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or frmMKTTRN0009.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            strSql = "Select PO_No, A.form_type, ISNULL(B.Form_No,'') form_no, " & " CASE ISNULL(PO_NO,'-1') " & " WHEN '-1' THEN 0 " & " Else '1' " & " End as Checked" & " from Forms_mst as a " & " Left Outer Join Forms_DTL as B ON  a.unit_code=b.unit_code and A.Form_Type=B.Form_Type and PO_No='" & Trim(INVOICE_NO) & "' and Doc_Type='9999' and a.unit_code='" & gstrUNITID & "' WHERE USABILITY IN (1,3) "

        ElseIf frmMKTTRN0009.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then

            If Len(Trim(RefNo)) = 0 Then
                strSql = "Select form_type,'' as form_no, 0 as checked from Forms_mst WHERE unit_code='" & gstrUNITID & "' and USABILITY IN (1,3)"
            Else

                If Trim(InvType) = "REJECTION" Then
                    If Trim(RejType) = "GRN" Then
                        ' CASE REJECTION TYPE IS GRN
                        strPO = CStr(Val(Find_Value("Select PUR_ORDER_NO FROm GRN_HDR where unit_code='" & gstrUNITID & "' and DOC_No IN (" & Trim(RefNo) & ")")))
                        strSql = "Select PO_No, A.form_type, ISNULL(B.Form_No,'') form_no, " & " CASE ISNULL(PO_NO,'-1') " & " WHEN '-1' THEN 0 " & " Else '1' " & " End as Checked " & " from Forms_mst as a " & " Left Outer Join Forms_DTL as B ON  A.UNIT_CODE=B.UNIT_CODE and A.Form_Type=B.Form_Type and PO_No='" & strPO & "' and Doc_Type='99' and Serial_No=0 " & " WHERE USABILITY IN (1,3) and A.UNIT_CODE='" & gstrUNITID & "' "
                    Else
                        ' CASE REJECTION TYPE IS LRN
                        strSql = "Select form_type,'' as form_no, 0 as checked from Forms_mst WHERE unit_code='" & gstrUNITID & "' and USABILITY IN (1,3)"
                    End If
                Else
                    ' IF NOT REJECTION TYPE TAKE UP FORM DETAIL FROM SALES ORDER ENTRY/
                    strSql = "Select PO_No, A.form_type, ISNULL(B.Form_No,'') form_no, " & " CASE ISNULL(PO_NO,'-1') " & " WHEN '-1' THEN 0 " & " Else '1' " & " End as Checked " & " from Forms_mst as a " & " Left Outer Join Forms_DTL as B ON A.UNIT_CODE=B.UNIT_CODE AND B.Account_code='" & Trim(Cust_Code) & "' and  A.Form_Type=B.Form_Type and PO_No='" & Trim(RefNo) & "' and Doc_Type='9998' " & " WHERE USABILITY IN (1,3) and a.unit_code='" & gstrUNITID & "' "
                End If
            End If
        End If

        rsForm.GetResult(strSql)
        fpForms.MaxRows = 0
        If rsForm.RowCount > 0 Then

            intRow = 1
            fpForms.MaxRows = rsForm.RowCount

            Do While Not rsForm.EOFRecord
                fpForms.Row = intRow
                fpForms.Col = 2

                fpForms.Text = rsForm.GetValue("form_type") : fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                fpForms.Col = 3

                fpForms.Text = rsForm.GetValue("form_no") : fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

                If rsForm.GetValue("checked") = 1 Then
                    fpForms.Col = 1
                    fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    fpForms.Value = CheckState.Checked
                Else
                    With fpForms
                        .Col = 1
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                        .Value = CheckState.Unchecked
                    End With
                End If
                rsForm.MoveNext()
                intRow = intRow + 1
            Loop

        End If

        Dim arrMain() As String
        Dim arrDet() As String
        Dim intOuterCount As Short
        Dim intInnerCount As Short
        With fpForms
            If Len(InvoiceFormDetails) > 0 Then
                arrMain = Split(InvoiceFormDetails, "^")

                For intOuterCount = 0 To UBound(arrMain) - 1
                    arrDet = Split(arrMain(intOuterCount), "|")
                    For intInnerCount = 1 To .MaxRows
                        .Row = intInnerCount
                        .Col = 2
                        If Trim(arrDet(0)) = Trim(.Text) Then
                            .Col = 1
                            .Value = CStr(1)

                            .Col = 3
                            .Value = Trim(arrDet(1))
                        End If
                    Next intInnerCount
                Next
            End If
        End With

        rsForm.ResultSetClose()

        rsForm = Nothing

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Public Sub fillGrid()
        Dim strSql As String
        Dim intRow As Short
        Dim rsForm As New ADODB.Recordset

        Dim strSONO As String
        Dim rsformName As New ClsResultSetDB

        If frmMKTTRN0001.cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or frmMKTTRN0001.cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If Len(frmMKTTRN0001.txtReferenceNo.Text) <= 0 Then
                Exit Sub
            End If
            strSONO = Trim(frmMKTTRN0001.txtReferenceNo.Text)
            strSql = " Select Form_type,form_no,1 as checked from forms_dtl where unit_code='" & gstrUNITID & "' and Account_code='" & Trim(Cust_Code) & "' and doc_type=9998 and po_no='" & Trim(frmMKTTRN0001.txtReferenceNo.Text) & "'" & " And Amendment_no ='" & Trim(frmMKTTRN0001.txtAmendmentNo.Text) & "'" & " Union All select form_type,'',0 from forms_mst where unit_code='" & gstrUNITID & "' and USABILITY IN (1,3) and form_type not in  (" & " select form_type from forms_dtl where unit_code='" & gstrUNITID & "' and Account_code='" & Trim(Cust_Code) & "' and doc_type=9998 and po_no='" & Trim(frmMKTTRN0001.txtReferenceNo.Text) & "'" & " And Amendment_no ='" & Trim(frmMKTTRN0001.txtAmendmentNo.Text) & "')"
        End If

        If frmMKTTRN0001.cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strSql = "select form_type,'' form_no,0 as checked from forms_mst WHERE unit_code='" & gstrUNITID & "' and USABILITY IN (1,3)"
        End If

        rsForm.Open(strSql, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        fpForms.MaxRows = 0

        If Not rsForm.EOF And Not rsForm.BOF Then
            rsForm.MoveFirst()

            intRow = 1
            fpForms.MaxRows = rsForm.RecordCount

            Do While Not rsForm.EOF
                fpForms.Row = intRow
                fpForms.Col = 2
                fpForms.Text = rsForm.Fields("form_type").Value : fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                fpForms.Col = 3
                fpForms.Text = rsForm.Fields("form_no").Value : fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                If rsForm.Fields("checked").Value = 1 Then
                    fpForms.Col = 1
                    fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    fpForms.Value = CheckState.Checked

                    fpForms.Col = 3
                    fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                Else
                    With fpForms
                        .Col = 1
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                        .Value = CheckState.Unchecked
                    End With
                End If
                If Not rsForm.EOF Then
                    rsForm.MoveNext()
                End If
                intRow = intRow + 1
            Loop

        End If
        Dim arrMain() As String
        Dim arrDet() As String
        Dim intOuterCount As Short
        Dim intInnerCount As Short
        With fpForms
            If Len(InvoiceFormDetails) > 0 Then
                arrMain = Split(InvoiceFormDetails, "^")
                For intOuterCount = 0 To UBound(arrMain) - 1
                    arrDet = Split(arrMain(intOuterCount), "|")
                    For intInnerCount = 1 To .MaxRows
                        .Row = intInnerCount
                        .Col = 2
                        If Trim(arrDet(0)) = Trim(.Text) Then
                            .Col = 1
                            .Value = CStr(1)
                            .Col = 3

                            .Col = 3
                            .Value = Trim(arrDet(1))
                            fpForms.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                        End If
                    Next intInnerCount
                Next
            End If
        End With

    End Sub

    Private Sub frmMKTTRN0041_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub



    Public Property ParentForm_Renamed() As String
        Get
            ParentForm_Renamed = mstrParentForm
        End Get
        Set(ByVal Value As String)
            mstrParentForm = Value
        End Set
    End Property


    Public Property Cust_Code() As String
        Get
            Cust_Code = mstrCust_code
        End Get
        Set(ByVal Value As String)
            mstrCust_code = Value
        End Set
    End Property

    
    Private Sub chkall_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkall.Click

        With fpForms
            If chkall.CheckState = System.Windows.Forms.CheckState.Checked Then
                .Row = -1
                .Col = 1
                .Value = System.Windows.Forms.CheckState.Checked
            Else
                .Row = -1
                .Col = 1
                .Value = System.Windows.Forms.CheckState.Unchecked
            End If
        End With

    End Sub

End Class