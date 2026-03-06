Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0010
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0010.frm
	' Function          :   Used to add sales TERMS
	' Created By        :   Nisha
	' Created On        :   15 May, 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 1
	'11/03/02 Validate for decimalaentry in Credit Days. on form no 4054
	'===================================================================================
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    25/04/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------
    'Dim m_strsql As String
    Dim rstAD As New ClsResultSetDB
    Dim mboolValid As Boolean
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Dim Index As Short = cmdOK.GetIndex(eventSender)
        Select Case Index
            Case 0
                If ValidRecord() = False Then
                    m_strSpecialNotes = txtSpecialNotes.Text
                    m_strPaymentTerms = txtPaymentTerms.Text
                    m_strPricesAre = txtPricesAre.Text
                    m_strPkgAndFwd = txtPkg.Text
                    m_strFreight = txtFreight.Text
                    m_strTransitInsurance = txtTransitInsurance.Text
                    m_strOctroi = txtOctroi.Text
                    m_strModeOfDespatch = txtModeOfDespatch.Text
                    m_strDeliverySchedule = txtDeliverySchedule.Text
                    Me.Dispose()
                Else
                    Exit Sub
                End If
        End Select
    End Sub

    Private Sub frmMKTTRN0010_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        mboolValid = False
    End Sub

    Private Sub txtDeliverySchedule_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDeliverySchedule.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDeliverySchedule_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDeliverySchedule.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtDeliverySchedule_Validating(txtDeliverySchedule, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtSpecialNotes.Focus()
            End If
        End If
    End Sub

    Private Sub txtDeliverySchedule_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDeliverySchedule.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim rssalesTerms As ClsResultSetDB
        Dim strsalesTerms As String
        mboolValid = False
        If Len(Trim(txtDeliverySchedule.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and  SaleTerms_Type ='DL' and Description ='" & Trim(txtDeliverySchedule.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            Else
                cmdOK(0).Focus()
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtFreight_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFreight.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtFreight_Validating(txtFreight, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtTransitInsurance.Focus()
            End If
        End If
    End Sub

    Private Sub txtFreight_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFreight.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFreight_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFreight.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        mboolValid = False
        If Len(Trim(txtFreight.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='FR' and Description ='" & Trim(txtFreight.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtModeOfDespatch_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtModeOfDespatch.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtModeOfDespatch_Validating(txtModeOfDespatch, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtDeliverySchedule.Focus()
            End If
        End If
    End Sub

    Private Sub txtModeOfDespatch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtModeOfDespatch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtModeOfDespatch_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtModeOfDespatch.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim rssalesTerms As ClsResultSetDB
        Dim strsalesTerms As String
        mboolValid = False
        If Len(Trim(txtModeOfDespatch.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='MO' and Description ='" & Trim(txtModeOfDespatch.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtOctroi_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtOctroi.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtOctroi_Validating(txtOctroi, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtModeOfDespatch.Focus()
            End If
        End If

    End Sub

    Private Sub txtOctroi_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOctroi.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtOctroi_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtOctroi.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim rssalesTerms As ClsResultSetDB
        Dim strsalesTerms As String
        mboolValid = False
        If Len(Trim(txtOctroi.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='OC' and Description ='" & Trim(txtOctroi.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtPaymentTerms_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPaymentTerms.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtPaymentTerms_Validating(txtPaymentTerms, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtPricesAre.Focus()
            End If
        End If

    End Sub

    Private Sub txtPaymentTerms_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPaymentTerms.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPaymentTerms_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtPaymentTerms.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        mboolValid = False
        If Len(Trim(txtPaymentTerms.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='PY' and Description ='" & Trim(txtPaymentTerms.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtPkg_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPkg.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtPkg_Validating(txtPkg, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtFreight.Focus()
            End If
        End If
    End Sub

    Private Sub txtPkg_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPkg.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPkg_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtPkg.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        mboolValid = False
        If Len(Trim(txtPkg.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='PK' and Description ='" & Trim(txtPkg.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtPricesAre_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPricesAre.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtPricesAre_Validating(txtPricesAre, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtPkg.Focus()
            End If
        End If
    End Sub

    Private Sub txtPricesAre_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPricesAre.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPricesAre_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtPricesAre.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        mboolValid = False
        If Len(Trim(txtPricesAre.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='PR' and Description ='" & Trim(txtPricesAre.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSpecialNotes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSpecialNotes.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            cmdOK(0).Focus()
        End If
    End Sub

    Private Sub txtSpecialNotes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSpecialNotes.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTransitInsurance_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtTransitInsurance.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtTransitInsurance_Validating(txtTransitInsurance, New System.ComponentModel.CancelEventArgs(False))
            If mboolValid = False Then
                txtOctroi.Focus()
            End If
        End If

    End Sub

    Private Sub txtTransitInsurance_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTransitInsurance.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTransitInsurance_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtTransitInsurance.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        mboolValid = False
        If Len(Trim(txtTransitInsurance.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            mboolValid = True
            GoTo EventExitSub
        Else
            strsalesTerms = "Select Description From SaleTerms_Mst Where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='TR' and Description ='" & Trim(txtTransitInsurance.Text) & "'"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                mboolValid = True
                GoTo EventExitSub
            End If
        End If
        mboolValid = False
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Public Sub addvaluestoList(ByRef mstrsalesType As String, ByRef mcmbBox As System.Windows.Forms.ComboBox, ByRef mode As String)
        Dim strSalesMst As String
        Dim rssalesMst As ClsResultSetDB
        Dim maxItems As Short
        Dim currentloopcount As Short
        Select Case mode
            Case "MODE_ADD"
                strSalesMst = "Select Description from SaleTerms_Mst where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='" & mstrsalesType & "'"
            Case "MODE_EDIT", "MODE_VEIW"
                strSalesMst = "Select Description from SaleTerms_Mst where unit_code ='" & gstrUNITID & "' and SaleTerms_Type ='" & mstrsalesType & "' AND DESCRIPTION <> '"
                strSalesMst = strSalesMst & ObsoleteManagement.GetItemString(mcmbBox, 0) & "'"
        End Select
        rssalesMst = New ClsResultSetDB
        rssalesMst.GetResult(strSalesMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        maxItems = rssalesMst.GetNoRows
        rssalesMst.MoveFirst()
        Select Case mode
            Case "MODE_ADD"
                If maxItems >= 1 Then
                    For currentloopcount = 0 To maxItems - 1
                        Call mcmbBox.Items.Insert(currentloopcount, rssalesMst.GetValue("Description"))
                        rssalesMst.MoveNext()
                    Next
                End If
                mcmbBox.SelectedIndex = 0
            Case "MODE_EDIT", "MODE_VEIW"
                If maxItems >= 1 Then
                    For currentloopcount = 1 To maxItems
                        Call mcmbBox.Items.Insert(currentloopcount, rssalesMst.GetValue("Description"))
                        rssalesMst.MoveNext()
                    Next
                End If
                mcmbBox.SelectedIndex = 0
        End Select
    End Sub
    Public Function formload(ByRef mpstrmode As String) As Object
        Select Case mpstrmode
            Case "MODE_ADD"
                txtPaymentTerms.Items.Clear()
                If Len(Trim(m_strPaymentTerms)) > 0 Then
                    txtPaymentTerms.Items.Insert(0, m_strPaymentTerms)
                    Call addvaluestoList("PY", txtPaymentTerms, "MODE_EDIT")
                Else
                    Call addvaluestoList("PY", txtPaymentTerms, mpstrmode)
                End If
                txtPricesAre.Items.Clear()
                If Len(Trim(m_strPricesAre)) > 0 Then
                    txtPricesAre.Items.Insert(0, m_strPricesAre)
                    Call addvaluestoList("PR", txtPricesAre, "MODE_EDIT")
                Else
                    Call addvaluestoList("PR", txtPricesAre, mpstrmode)
                End If
                txtPkg.Items.Clear()
                If Len(Trim(m_strPkgAndFwd)) > 0 Then
                    txtPkg.Items.Insert(0, m_strPkgAndFwd)
                    Call addvaluestoList("PK", txtPkg, "MODE_EDIT")
                Else
                    Call addvaluestoList("PK", txtPkg, mpstrmode)
                End If
                txtFreight.Items.Clear()
                If Len(Trim(m_strFreight)) > 0 Then
                    txtFreight.Items.Insert(0, m_strFreight)
                    Call addvaluestoList("FR", txtFreight, "MODE_EDIT")
                Else
                    Call addvaluestoList("FR", txtFreight, mpstrmode)
                End If
                txtTransitInsurance.Items.Clear()
                If Len(Trim(m_strTransitInsurance)) > 0 Then
                    txtTransitInsurance.Items.Insert(0, m_strTransitInsurance)
                    Call addvaluestoList("TR", txtTransitInsurance, "MODE_EDIT")
                Else
                    Call addvaluestoList("TR", txtTransitInsurance, mpstrmode)
                End If
                txtOctroi.Items.Clear()
                If Len(Trim(m_strOctroi)) > 0 Then
                    txtOctroi.Items.Insert(0, m_strOctroi)
                    Call addvaluestoList("OC", txtOctroi, "MODE_EDIT")
                Else
                    Call addvaluestoList("OC", txtOctroi, mpstrmode)
                End If
                txtModeOfDespatch.Items.Clear()
                If Len(Trim(m_strModeOfDespatch)) > 0 Then
                    txtModeOfDespatch.Items.Insert(0, m_strModeOfDespatch)
                    Call addvaluestoList("MO", txtModeOfDespatch, "MODE_EDIT")
                Else
                    Call addvaluestoList("MO", txtModeOfDespatch, mpstrmode)
                End If
                txtDeliverySchedule.Items.Clear()
                If Len(Trim(m_strDeliverySchedule)) > 0 Then
                    txtDeliverySchedule.Items.Insert(0, m_strDeliverySchedule)
                    Call addvaluestoList("DL", txtDeliverySchedule, "MODE_EDIT")
                Else
                    Call addvaluestoList("DL", txtDeliverySchedule, mpstrmode)
                End If
            Case "MODE_VEIW", "MODE_EDIT"
                rstAD.GetResult(m_pstrSql)
                If rstAD.GetNoRows > 0 Then
                    txtSpecialNotes.Text = rstAD.GetValue("Special_Remarks")
                    txtPaymentTerms.Items.Clear()
                    txtPaymentTerms.Items.Insert(0, rstAD.GetValue("Pay_Remarks"))
                    Call addvaluestoList("PY", txtPaymentTerms, mpstrmode)
                    txtPricesAre.Items.Clear()
                    txtPricesAre.Items.Insert(0, rstAD.GetValue("Price_Remarks"))
                    Call addvaluestoList("PR", txtPricesAre, mpstrmode)
                    txtPkg.Items.Clear()
                    txtPkg.Items.Insert(0, rstAD.GetValue("Packing_Remarks"))
                    Call addvaluestoList("PK", txtPkg, mpstrmode)
                    txtFreight.Items.Clear()
                    txtFreight.Items.Insert(0, rstAD.GetValue("Frieght_Remarks"))
                    Call addvaluestoList("FR", txtFreight, mpstrmode)
                    txtTransitInsurance.Items.Clear()
                    txtTransitInsurance.Items.Insert(0, rstAD.GetValue("Transport_Remarks"))
                    Call addvaluestoList("TR", txtTransitInsurance, mpstrmode)
                    txtOctroi.Items.Clear()
                    txtOctroi.Items.Insert(0, rstAD.GetValue("Octorai_Remarks"))
                    Call addvaluestoList("OC", txtOctroi, mpstrmode)
                    txtModeOfDespatch.Items.Clear()
                    txtModeOfDespatch.Items.Insert(0, rstAD.GetValue("Mode_Despatch"))
                    Call addvaluestoList("MO", txtModeOfDespatch, mpstrmode)
                    txtDeliverySchedule.Items.Clear()
                    txtDeliverySchedule.Items.Insert(0, rstAD.GetValue("Delivery"))
                    Call addvaluestoList("DL", txtDeliverySchedule, mpstrmode)
                End If
        End Select
    End Function
    Private Function ValidRecord() As Boolean
        Dim blnInvalidData As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lno As Integer
        On Error GoTo Err_Handler
        ValidRecord = False
        lno = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(txtPaymentTerms.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".PaymentTerms"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtPaymentTerms
        End If
        If Len(Trim(txtPricesAre.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".PricesAre"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtPricesAre
        End If
        If Len(Trim(txtPkg.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".Pkg. & Forwarding"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtPkg
        End If
        If Len(Trim(txtFreight.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".Freight"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtFreight
        End If
        If Len(Trim(txtTransitInsurance.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".TransitInsurance"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtTransitInsurance
        End If
        If Len(Trim(txtOctroi.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".Octroi"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtOctroi
        End If
        If Len(Trim(txtModeOfDespatch.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".ModeOFDespatch"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtModeOfDespatch
        End If

        If Len(Trim(txtDeliverySchedule.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lno & ".DeliverySchedule"
            lno = lno + 1
            If ctlBlank Is Nothing Then ctlBlank = txtDeliverySchedule
        End If
0:      strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        lno = lno + 1
        If blnInvalidData = True Then
            ValidRecord = blnInvalidData
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            ctlBlank.Focus()
            Exit Function
        End If
        ValidRecord = blnInvalidData
        gblnCancelUnload = True
        gblnFormAddEdit = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
End Class