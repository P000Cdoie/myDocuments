Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0009a
	Inherits System.Windows.Forms.Form
	' -----------------------------------------------------------------------
	' Copyright (c)   :MIND Ltd.
	' Form Name       :FrmMKTTRN0009a
	' Function Name   :List of LRN / GRN Documents FOR REJECTION INVOICE TRACKING.
	' Created By      :Sandeep Chadha
	' Created On      :31-March-2005
	' Modify Date     :NIL
	' Revision History:-
	' -----------------------------------------------------------------------
    'Revised By      : Manoj Vaish
    'Revision On     : 16 dec 2008
    'Issue ID        : eMpro-20081216-24902
    'History         : To configure the Vendor Rejection Invoice from GRIN/LRN
    '***********************************************************************************
    'Revised By      : Manoj Vaish
    'Revision On     : 06 May 2009
    'Issue ID        : eMpro-20090506-31083
    'History         : Rejection Invoice against LRN/GRN without batch tracking for MAPL
    '***********************************************************************************
    '=====================================================================
    'Revised By     : Amit Kumar (0670)
    'Revision Date  : 30 May 2011
    'Remarks        : Changes Done To Support Multiunit Function
    '=====================================================================
    ''Revised By:       Saurav Kumar
    ''Revised On:       04 Oct 2013
    ''Issue ID  :       10462231 - eMpro ISuite Changes
    '***********************************************************************************************************************************
    ''Revised By:       Geetanjali Aggrawal
    ''Revised On:       06-Mar-2013
    ''Purpose   :       form frmMKTTRN0009_HILEX added for HILEX multi unit
    '***********************************************************************************************************************************

    Dim mstrRejectionType As String
    Dim mColHeaderRef As System.Windows.Forms.ColumnHeader
    Dim mstrItemText As String
    Dim mstrVendor_code As String
    Dim Intcounter As Short
    Dim ACTIDX As Short
    Dim bool_Item_Check As Boolean = False
    Dim mblnBatchTracking As Boolean = False

    Private Function GetallSelectedItems() As String
        ' Author        : Sandeep Chadha
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : RETURN THE LIST OF SELECTED DOC NO's
        ' Datetime      : 31-March-2005
        '------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim intPoNo As Integer
        mstrItemText = ""
        For intRow = 0 To lvwDocumentDetails.Items.Count - 1
            With lvwDocumentDetails
                If .Items.Item(intRow).Checked = True Then
                    If Len(Trim(mstrItemText)) = 0 Then
                        mstrItemText = .Items.Item(intRow).Text
                    Else
                        mstrItemText = mstrItemText & "," & .Items.Item(intRow).Text
                    End If
                End If
            End With
        Next
        GetallSelectedItems = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Sub SetHeaders()
        Dim ColHeaderRef As Object
        On Error GoTo ErrHandler ' Error Handler
        Dim intWidth As Short
        With lvwDocumentDetails
            If mstrRejectionType = "GRN" Then
                intWidth = 120
            Else
                intWidth = 100
            End If
            .CheckBoxes = True
            ColHeaderRef = .Columns.Add("")
            ColHeaderRef.Text = "Doc No"
            ColHeaderRef.Width = intWidth
            ColHeaderRef = .Columns.Add("")
            ColHeaderRef.Text = "Doc_No"
            ColHeaderRef.Width = 0
            ColHeaderRef = .Columns.Add("")
            ColHeaderRef.Text = "Date"
            ColHeaderRef.Width = intWidth
            ColHeaderRef = .Columns.Add("")
            ColHeaderRef.Text = "GRN"
            If mstrRejectionType = "GRN" Or mblnBatchTracking = False Then
                ColHeaderRef.Width = 0
            Else
                ColHeaderRef.Width = intWidth
            End If
            ColHeaderRef = .Columns.Add("")
            ColHeaderRef.Text = "PO No"
            If mstrRejectionType = "LRN" And mblnBatchTracking = False Then
                ColHeaderRef.Width = 0
            Else
                ColHeaderRef.Width = intWidth
            End If
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
    Sub ShowListofDocuments()
        Dim intItem As Object
        On Error GoTo ErrHandler 'Error Handler
        Dim strSql As String
        Dim rsTmp As New ClsResultSetDB
        Dim lstItemRef As System.Windows.Forms.ListViewItem
        Dim lstsubItemRef As System.Windows.Forms.ListViewItem.ListViewSubItem
        If RejectionType = "GRN" Then
            strSql = "SELECT REJINVOICE_GRIN_PVKNOCKING FROM SALES_PARAMETER WHERE UNIT_CODE= '" & gstrUNITID & "'"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) Then
                strSql = "Select Distinct GRIN_NO as Doc_No, GRN_DATE as [Date], GRIN_NO, PUR_ORDER_NO as PONo  From VW_INVREJ_GRIN_DETAIL_PVKNOCKED WHERE UNIT_CODE='" + gstrUNITID + "' AND Vendor_code='" & Vendor_code & "'"
            Else
                strSql = "Select Distinct GRIN_NO as Doc_No, GRN_DATE as [Date], GRIN_NO, PUR_ORDER_NO as PONo  From vw_INVREJ_GRIN_DETAIL WHERE UNIT_CODE='" + gstrUNITID + "' AND Vendor_code='" & Vendor_code & "'"
            End If

        Else
            If mblnBatchTracking = True Then
                'Added By ekta uniyal 
                ' strSql = "Select Distinct Doc_No as Doc_No,[Date], GRIN_NO, PUR_ORDER_NO as PONo From vw_INVREJ_LRN_DETAIL " & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Vendor_code='" & Vendor_code & "'"
                If GetPlantName() = "HILEX" Then
                    strSql = "Select Distinct Doc_No as Doc_No,[Date], GRIN_NO, PUR_ORDER_NO as PONo From vw_INVREJ_LRN_DETAIL_HILEX " & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Vendor_code='" & Vendor_code & "'"
                Else
                    strSql = "Select Distinct Doc_No as Doc_No,[Date], GRIN_NO, PUR_ORDER_NO as PONo From vw_INVREJ_LRN_DETAIL " & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Vendor_code='" & Vendor_code & "'"
                End If
            Else
                strSql = "Select Distinct Doc_No as Doc_No,[Date], GRIN_NO, PUR_ORDER_NO as PONo From VW_INVREJ_LRN_DETAIL_WB " & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Vendor_code='" & Vendor_code & "'"
            End If
        End If
        rsTmp.GetResult(strSql)
        intItem = 1
        Do While Not rsTmp.EOFRecord
            With lvwDocumentDetails
                lstItemRef = .Items.Add(rsTmp.GetValue("Doc_No"))
                lstItemRef.Text = rsTmp.GetValue("Doc_No")
                lstsubItemRef = lstItemRef.SubItems.Add(rsTmp.GetValue("Doc_No"))
                lstsubItemRef.Text = rsTmp.GetValue("Doc_No")
                lstsubItemRef = lstItemRef.SubItems.Add(VB6.Format(rsTmp.GetValue("Date"), gstrDateFormat))
                lstsubItemRef.Text = VB6.Format(rsTmp.GetValue("Date"), gstrDateFormat)
                lstsubItemRef = lstItemRef.SubItems.Add(rsTmp.GetValue("GRIN_NO"))
                lstsubItemRef.Text = rsTmp.GetValue("GRIN_NO")
                lstsubItemRef = lstItemRef.SubItems.Add(rsTmp.GetValue("PONo"))
                lstsubItemRef.Text = rsTmp.GetValue("PONo")
                rsTmp.MoveNext()
            End With
        Loop
        rsTmp.ResultSetClose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOk.Click
        On Error GoTo ErrHandler
        'Added by Geetanjali for HILEX multi unit
        If GetPlantName() = "HILEX" Then
            frmMKTTRN0009_HILEX.SelectedItems = GetallSelectedItems()
        Else
            frmMKTTRN0009.SelectedItems = GetallSelectedItems()
        End If
        If Len(mstrItemText) = 0 Then
            Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Exit Sub
        End If
        Me.Close()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub CmdSearch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSearch.Click
        txtSearchText_KeyPress(txtSearchText, New System.Windows.Forms.KeyPressEventArgs(Chr((13))))
    End Sub

    Private Sub frmMKTTRN0009a_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler ' Error Handler
        SetBackGroundColorNew(Me, True)
        Select Case RejectionType
            Case "LRN"
                lblHeader.Text = "LINE OF REJECTION"
                If mblnBatchTracking = False Then
                    optGRIN.Visible = False
                    optPONo.Visible = False
                Else
                    optGRIN.Visible = True
                    optPONo.Visible = True
                End If
            Case "GRN"
                lblHeader.Text = "GRN"
        End Select
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / 2) - VB6.PixelsToTwipsY(Me.Height) / 2)
        Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 2) - VB6.PixelsToTwipsX(Me.Width) / 2)
        SetHeaders()
        ShowListofDocuments()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0009a_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
    Public Property Batch_Tracking() As Boolean
        Get
            Batch_Tracking = mblnBatchTracking
        End Get
        Set(ByVal value As Boolean)
            mblnBatchTracking = value
        End Set
    End Property
    Private Sub lvwDocumentDetails_AfterLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.LabelEditEventArgs) Handles lvwDocumentDetails.AfterLabelEdit
        Dim Cancel As Boolean = eventArgs.CancelEdit
        Dim NewString As String = eventArgs.Label
        'Edti is not allowed
        Cancel = True
    End Sub
    Private Sub lvwDocumentDetails_BeforeLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.LabelEditEventArgs) Handles lvwDocumentDetails.BeforeLabelEdit
        Dim Cancel As Boolean = eventArgs.CancelEdit
        Cancel = True
    End Sub
    Private Sub optDate_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDate.CheckedChanged
        If eventSender.Checked Then
            Call SortGrid(1)
            txtSearchText.Focus()
        End If
    End Sub
    Private Sub optGRIN_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGRIN.CheckedChanged
        If eventSender.Checked Then
            Call SortGrid(2)
            txtSearchText.Focus()
        End If
    End Sub
    Private Sub optPONo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPONo.CheckedChanged
        If eventSender.Checked Then
            Call SortGrid(3)
            txtSearchText.Focus()
        End If
    End Sub
    Private Sub optSearchDoc_No_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchDoc_No.CheckedChanged
        If eventSender.Checked Then
            Call SortGrid(0)
            txtSearchText.Focus()
        End If
    End Sub
    Sub SortGrid(ByRef Index As Short)
        On Error GoTo ErrHandler
        With lvwDocumentDetails
            .Sort()
            ListViewColumnSorter.SortListView(lvwDocumentDetails, Index, SortOrder.Ascending)
            .Sorting = System.Windows.Forms.SortOrder.Ascending
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSearchText_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSearchText.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Err_Handler
        Dim ListItem As System.Windows.Forms.ListViewItem

        If optSearchDoc_No.Checked = True Then
            ListItem = SearchText(txtSearchText.Text & Chr(KeyAscii), optSearchDoc_No, lvwDocumentDetails)
        ElseIf optDate.Checked = True Then
            ListItem = SearchText(txtSearchText.Text & Chr(KeyAscii), optDate, lvwDocumentDetails, CStr(2))
        ElseIf optPONo.Checked = True Then
            ListItem = SearchText(txtSearchText.Text & Chr(KeyAscii), optPONo, lvwDocumentDetails, CStr(4))
        ElseIf optGRIN.Checked = True Then
            ListItem = SearchText(txtSearchText.Text & Chr(KeyAscii), optPONo, lvwDocumentDetails, CStr(3))
        End If
        If ListItem Is Nothing Then
            GoTo EventExitSub
        Else
            ListItem.EnsureVisible()
            ListItem.Selected = True
            ListItem.Font = VB6.FontChangeBold(ListItem.Font, True)
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub lvwDocumentDetails_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwDocumentDetails.ItemChecked
        Dim Item As System.Windows.Forms.ListViewItem = lvwDocumentDetails.Items(e.Item.Index)
        On Error GoTo ErrHandler
        Dim intSubItem As Short
        Dim Index As Object
        If bool_Item_Check = True Then
            Exit Sub
        End If
        MULTIPLESO = 0
        For Intcounter = 0 To Me.lvwDocumentDetails.Items.Count - 1
            If Me.lvwDocumentDetails.Items.Item(Intcounter).Checked = True Then
                MULTIPLESO = MULTIPLESO + 1
                Index = Intcounter
            End If
        Next Intcounter
        If ACTIDX = 0 Or Index = 0 Then
            ACTIDX = Index
        End If
        If MULTIPLESO > 1 Then
            MsgBox("More than one selection is not available", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Empower")
            Call GRIDSELECT()
            Exit Sub
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub GRIDSELECT()
        On Error GoTo ErrHandler
        For Intcounter = 0 To Me.lvwDocumentDetails.Items.Count - 1
            bool_Item_Check = True
            Me.lvwDocumentDetails.Items.Item(Intcounter).Checked = False
        Next Intcounter
        Me.lvwDocumentDetails.Items.Item(ACTIDX).Checked = True
        bool_Item_Check = False
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class