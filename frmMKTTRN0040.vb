Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0040
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0040.frm
	' Function          :   Used to select GRINS
	' Created By        :   Arshad Ali
	' Created On        :   18 Feb, 2005
    ' Revision History  :   Nisha Rai
    ' MODIFIED BY AJAY SHUKLA ON 11/MAY/2011 FOR MULTIUNIT CHANGE
	'===================================================================================
	Dim mCtlGRINNo As System.Windows.Forms.ColumnHeader
	Dim mCtlLocationTo As System.Windows.Forms.ColumnHeader
	Dim mCtlItemCode As System.Windows.Forms.ColumnHeader
	Dim mCtlDescription As System.Windows.Forms.ColumnHeader
	Dim mCtlQuantity As System.Windows.Forms.ColumnHeader
	
	Dim intCheckCounter As Short
	Dim mListItemUserId As System.Windows.Forms.ListViewItem
	Dim mstrInvType As String
	Dim mstrInvSubType As String
	Dim mstrGRINText As String
	Dim blnExpinv As Boolean
	Dim intIteminSp As Short
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrHandler
		Me.Close()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		'Code Modified By   -   Arshad Ali
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		On Error GoTo ErrHandler
		mstrGRINText = ""
		Dim intSubItem As Short
		'Sort according to Item code
		With lvwGRIN
            .Sort()

            ListViewColumnSorter.SortListView(lvwGRIN, 3, SortOrder.Ascending)
        End With

        For intSubItem = 0 To lvwGRIN.Items.Count - 1
            If Me.lvwGRIN.Items.Item(intSubItem).Checked = True Then
                mstrGRINText = mstrGRINText & Trim(Me.lvwGRIN.Items.Item(intSubItem).SubItems(1).Text) & "|" & Trim(Me.lvwGRIN.Items.Item(intSubItem).SubItems(2).Text) & "|" & Trim(Me.lvwGRIN.Items.Item(intSubItem).SubItems(3).Text) & "|" & Trim(Me.lvwGRIN.Items.Item(intSubItem).SubItems(4).Text) & "|" & Trim(Me.lvwGRIN.Items.Item(intSubItem).SubItems(5).Text) & "|" & Trim(Me.lvwGRIN.Items.Item(intSubItem).SubItems(6).Text) & "|" & Trim(Me.lvwGRIN.Items.Item(intSubItem).SubItems(7).Text) & ","
            End If
        Next intSubItem
        If Len(mstrGRINText) = 0 Then
            MsgBox("Select Atleast one GRIN and Item Code.", MsgBoxStyle.Information, "eMPro")
            Me.lvwGRIN.Focus()
            Exit Sub
        End If
        Me.Close()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0040_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Call AddColumnsInListView()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optGrin.Checked = True
        lvwGRIN.FullRowSelect = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddColumnsInListView()
        '***********************************
        'To add Columns Headers in the ListView in the form load
        '***********************************
        On Error GoTo ErrHandler
        With Me.lvwGRIN

            mCtlGRINNo = .Columns.Add("")
            mCtlGRINNo.Text = "GRIN No"
            mCtlGRINNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwGRIN.Width) / 4)


            mCtlGRINNo = .Columns.Add("")
            mCtlGRINNo.Text = "GRIN No"
            mCtlGRINNo.Width = 0

            mCtlLocationTo = .Columns.Add("")
            mCtlLocationTo.Text = "To Location"
            mCtlLocationTo.Width = 0 'lvwGRIN.Width / 4

            mCtlItemCode = .Columns.Add("")
            mCtlItemCode.Text = "Item Code"
            mCtlItemCode.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwGRIN.Width) / 4))

            mCtlDescription = .Columns.Add("")
            mCtlDescription.Text = "Description"
            mCtlDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwGRIN.Width) / 4))

            mCtlQuantity = .Columns.Add("")
            mCtlQuantity.Text = "Quantity"
            mCtlQuantity.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwGRIN.Width) / 4) - 200)

            mCtlQuantity = .Columns.Add("")
            mCtlQuantity.Text = "Rate"
            mCtlQuantity.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwGRIN.Width) / 4))

            mCtlQuantity = .Columns.Add("")
            mCtlQuantity.Text = "UOM"
            mCtlQuantity.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwGRIN.Width) / 4))

        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0040_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Me.Dispose()
    End Sub
    Private Sub lvwGRIN_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvwGRIN.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                CmdOk.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub SearchItem()
        '---------------------------------------------------------------------
        'Created By     -   Arshad Ali
        '---------------------------------------------------------------------
        Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
        On Error GoTo ErrHandler
        If optGrin.Checked Then
            itmFound = SearchText((txtsearch.Text), optGrin, lvwGRIN, "1")
        ElseIf optItem.Checked Then
            itmFound = SearchText((txtsearch.Text), optItem, lvwGRIN, "3")
        ElseIf optDescription.Checked Then
            itmFound = SearchText((txtsearch.Text), optDescription, lvwGRIN, "4")
        End If
        If itmFound Is Nothing Then ' If no match,
            Exit Sub
        Else

            itmFound.EnsureVisible() ' Scroll ListView to show found ListItem.
            itmFound.Selected = True ' Select the ListItem.
            ' Return focus to the control to see selection.
            lvwGRIN.Enabled = True
            If Len(txtsearch.Text) > 0 Then itmFound.Font = VB6.FontChangeBold(itmFound.Font, True)
        End If

        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub optDescription_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDescription.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwGRIN
                .Sort()
                ListViewColumnSorter.SortListView(lvwGRIN, 4, SortOrder.Ascending)
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub optItem_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItem.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwGRIN
                .Sort()
                ListViewColumnSorter.SortListView(lvwGRIN, 3, SortOrder.Ascending)
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub optGRIN_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGrin.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwGRIN
                .Sort()
                ListViewColumnSorter.SortListView(lvwGRIN, 0, SortOrder.Ascending)
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
    End Sub

    Public Function SelectDatafromGRIN(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByVal pstrItemCode As String, ByVal pstrCustomerCode As String, Optional ByRef pstrGRINNotin As String = "") As Object
        On Error GoTo ErrHandler
        Dim rsgrin As ClsResultSetDB
        Dim strSQL As String
        Dim intCount, intRecordCount, intInnerCount As Short

        strSQL = "SELECT H.DOC_NO AS GRIN_NO, H.TO_LOCATION, D.ITEM_CODE, ITEM_MST.DESCRIPTION, D.ACCEPTED_QUANTITY-D.DESPATCH_QUANTITY AS QUANTITY, D.ITEM_RATE, ITEM_MST.PUR_MEASURE_CODE "
        strSQL = strSQL & " FROM GRN_HDR H INNER JOIN GRN_DTL D"
        strSQL = strSQL & " ON H.DOC_TYPE = D.DOC_TYPE AND H.DOC_NO = D.DOC_NO AND H.FROM_LOCATION = D.FROM_LOCATION AND H.UNIT_CODE = D.UNIT_CODE INNER JOIN ITEM_MST"
        strSQL = strSQL & " ON ITEM_MST.ITEM_CODE = D.ITEM_CODE AND ITEM_MST.UNIT_CODE = D.UNIT_CODE"
        strSQL = strSQL & " WHERE H.DOC_CATEGORY='Z' AND H.QA_AUTHORIZED_CODE IS NOT NULL"
        strSQL = strSQL & " AND H.VENDOR_CODE='" & pstrCustomerCode & "'"
        strSQL = strSQL & " AND D.ITEM_CODE IN (" & pstrItemCode & ")"
        strSQL = strSQL & " AND D.ACCEPTED_QUANTITY-D.DESPATCH_QUANTITY > 0"
        strSQL = strSQL & " AND H.UNIT_CODE='" & gstrUNITID & "'"
        strSQL = strSQL & " ORDER BY H.DOC_NO, D.ITEM_CODE"

        rsgrin = New ClsResultSetDB
        rsgrin.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsgrin.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rsgrin.MoveFirst() 'move to first record
            For intCount = 1 To intRecordCount

                mListItemUserId = Me.lvwGRIN.Items.Add(rsgrin.GetValue("GRIN_NO"))
                If mListItemUserId.SubItems.Count > 1 Then
                    mListItemUserId.SubItems(1).Text = rsgrin.GetValue("GRIN_NO")
                Else
                    mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsgrin.GetValue("GRIN_NO")))
                End If
                If mListItemUserId.SubItems.Count > 2 Then
                    mListItemUserId.SubItems(2).Text = rsgrin.GetValue("TO_LOCATION")
                Else
                    mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsgrin.GetValue("TO_LOCATION")))
                End If


                If mListItemUserId.SubItems.Count > 3 Then
                    mListItemUserId.SubItems(3).Text = rsgrin.GetValue("ITEM_CODE")
                Else
                    mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsgrin.GetValue("ITEM_CODE")))
                End If

                If mListItemUserId.SubItems.Count > 4 Then
                    mListItemUserId.SubItems(4).Text = rsgrin.GetValue("DESCRIPTION")
                Else
                    mListItemUserId.SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsgrin.GetValue("DESCRIPTION")))
                End If


                If mListItemUserId.SubItems.Count > 5 Then
                    mListItemUserId.SubItems(5).Text = rsgrin.GetValue("QUANTITY")
                Else
                    mListItemUserId.SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsgrin.GetValue("QUANTITY")))
                End If


                If mListItemUserId.SubItems.Count > 6 Then
                    mListItemUserId.SubItems(6).Text = rsgrin.GetValue("ITEM_RATE")
                Else
                    mListItemUserId.SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsgrin.GetValue("ITEM_RATE")))
                End If


                If mListItemUserId.SubItems.Count > 7 Then
                    mListItemUserId.SubItems(7).Text = rsgrin.GetValue("PUR_MEASURE_CODE")
                Else
                    mListItemUserId.SubItems.Insert(7, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsgrin.GetValue("PUR_MEASURE_CODE")))
                End If
                rsgrin.MoveNext() 'move to next record
            Next intCount
        End If
        rsgrin.ResultSetClose()

        rsgrin = Nothing
        'Selectting Previous selected values
        Dim strMain() As String
        Dim strDet() As String
        If Len(pstrGRINNotin) > 0 Then
            strMain = Split(pstrGRINNotin, ",")
            For intCount = 0 To UBound(strMain) - 1
                strDet = Split(strMain(intCount), "|")
                For intInnerCount = 0 To lvwGRIN.Items.Count - 1


                    If Trim(Me.lvwGRIN.Items.Item(intInnerCount).SubItems(1).Text) = strDet(0) And Trim(Me.lvwGRIN.Items.Item(intInnerCount).SubItems(3).Text) = strDet(2) Then

                        Me.lvwGRIN.Items.Item(intInnerCount).Checked = True
                    End If
                Next intInnerCount
            Next intCount
        End If
        Me.ShowDialog()

        SelectDatafromGRIN = mstrGRINText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
End Class