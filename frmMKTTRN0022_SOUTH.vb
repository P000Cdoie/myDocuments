Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0022_SOUTH
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0021.frm
	' Function          :   Used to select items
	' Created By        :   Nisha
	' Created On        :   15 May, 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 3
	'03/10/2001 MARKED CHECKED BY BCs  for jobwork invoice changed on version 6
	'09/10/2001 jobwork invoice changed on version 7 for set status to one in case of Schedule of Daily/Monthly
	'22/03/2002 INCREASED FEILD SIZE OF DELIVERY TERMS,PAYMENT TERMS, DESC OF GOODS, EPC DESCRIPTION
	'===================================================================================
    'Revised By         : Manoj Kr. Vaish
    'Revised on         : 23-Sep-2008 Issue ID:eMpro-20080923-21892
    'Revised For        : Changes has been reverted for Export Invoice entry through sales order
    '===================================================================================
    'Revised By         :   Manoj Vaish
    'Revision Date      :   05 Mar 2009
    'Issue ID           :   eMpro-20090227-27987
    'Revision History   :   Changes for commercial invoice at Mate Units
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    19/05/2011
    'Revision History  -    Changes for Multi Unit
    '=======================================================================================
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   13-Mar-2013
    'Issue ID           :   10354980   
    'Revision History   :   Woco migration changes
    '**********************************************************************************************************************

    Dim mCtlHdrItemCode As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrDrawingNo As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrDescription As System.Windows.Forms.ColumnHeader
    Dim intCheckCounter As Short
    Dim mListItemUserId As System.Windows.Forms.ListViewItem
    Dim mstrItemText As String
    Dim mstrMode As String
    Dim mstrDocumentDate As String 'For storing Document Date
    Dim mstrCurrencyID As String 'For Currency ID
    Public mstrInvSubType As String 'for invoice sub type
    Dim mstrCreditTermId As String
    Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
        On Error GoTo ErrHandler
        If mstrMode = "MODE_VIEW" Or mstrMode = "MODE_EDIT" Then
            strValues = ""
            strValues = AddValuestoString()
            Me.Close()
            Exit Sub
        End If
        Me.Close()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOk.Click
        On Error GoTo ErrHandler
        'Samiksha commodity changes
        If ValidatebeforeSave() = True Then
            strValues = ""
            strValues = AddValuestoString()
            Me.Close()
            Exit Sub
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0022_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2.2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 3.5)
        SetBackGroundColorNew(Me, True)
        Me.DTPExchangeDate.Value = GetServerDate()
        '_cmdHelp_3.Enabled = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0022_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        Dim StrHelpSql As String
        Dim StrCodeReturn As String
        Dim strTempDocDate As String
        Dim str_currency_code As Object
        Dim objshowExchange As ADODB.Recordset
        Dim rs_currecny_code As New ClsResultSetDB
        Dim tempstr As String
        On Error GoTo ErrHandler
        Select Case Index
            Case 3 'Currency
                StrCodeReturn = ShowList(1, 2000, , "Currency_code", "Description", " Currency_mst", "")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Currency code defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtCurrency.Text = StrCodeReturn
                    If Len(Trim(txtCurrency.Text)) > 0 Then
                        If Trim(mstrDocumentDate) <> "" Then
                            strTempDocDate = VB6.Format(mstrDocumentDate, gstrDateFormat)
                        Else
                            strTempDocDate = GetServerDate()
                        End If
                        tempstr = "SET DATEFORMAT 'mdy'" & vbCrLf & "SELECT CExch_MultiFactor From Gen_CurExchMaster Where unit_code='" & gstrUNITID & "' and CExch_CurrencyTo='" & Trim(txtCurrency.Text) & "' AND CExch_InOut=1 AND '" & getDateForDB(strTempDocDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo "
                        objshowExchange = New ADODB.Recordset
                        objshowExchange.Open(tempstr, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If Not objshowExchange.EOF Or Not objshowExchange.EOF Then
                            Me.txtExchangeRate.Text = objshowExchange.Fields("CExch_MultiFactor").Value
                            txtExchangeValue.Text = CStr(Val(objshowExchange.Fields("CExch_MultiFactor").Value))
                        Else
                            MsgBox("Exchange Rate Not Defined for the Current Date", MsgBoxStyle.Information, "empower")
                            Me.txtCurrency.Text = ""
                            Me.txtExchangeRate.Text = "1.00"
                            txtExchangeValue.Text = "1.00"
                        End If
                        objshowExchange.Close()
                        objshowExchange = Nothing
                    Else
                        Me.txtExchangeRate.Text = "1.00"
                        txtExchangeValue.Text = "1.00" '
                    End If
                End If
            Case 4 'Country Of Origin
                StrCodeReturn = ShowList(1, 2000, Trim(txtOrigin_Status.Text), "Country_des", "Country_c", " country_mst", "")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Country Of Origin defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtOrigin_Status.Text = StrCodeReturn
                End If
            Case 5 'Country Of Final Destination
                StrCodeReturn = ShowList(1, 2000, Trim(Me.txtCtryFinalDest.Text), " Country_des", "Country_c", " country_mst", "")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Country Of Final Destination defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtCtryFinalDest.Text = StrCodeReturn
                End If
            Case 6 'Pre Carriage by
                StrCodeReturn = ShowList(1, 2000, Trim(txtPreCarriage.Text), " key2 ", " Descr ", " lists", " and key1='precarriage_by'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Pre-Carriage by defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtPreCarriage.Text = StrCodeReturn
                End If
            Case 7 'Place of Receipt
                StrCodeReturn = ShowList(1, 2000, Trim(txtPlaceOfReceipt.Text), " key2 ", " Descr ", " lists", " and key1='PlaceOfReceipt'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Place Of Receipt defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtPlaceOfReceipt.Text = StrCodeReturn
                End If
            Case 8 'Port Of Loading
                StrCodeReturn = ShowList(1, 2000, Trim(txtPortOfLoading.Text), " key2 ", " Descr ", " lists", " and key1='loading port'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Port Of Loading defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtPortOfLoading.Text = StrCodeReturn
                End If
            Case 9 'Port of Discharge
                StrCodeReturn = ShowList(1, 2000, Trim(txtPortOfDischarge.Text), " key2 ", " Descr ", " lists", " and key1='discharge port'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Port Of Descharge defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtPortOfDischarge.Text = StrCodeReturn
                End If
            Case 10 'Final Destination
                StrCodeReturn = ShowList(1, 2000, Trim(txtFinalDest.Text), " key2 ", " Descr ", " lists", " and key1='final dest'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Final Destination defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtFinalDest.Text = StrCodeReturn
                End If
            Case 11 'Contract Type
                StrCodeReturn = ShowList(1, 2000, Trim(txtContract.Text), " key2 ", " Descr ", " lists", " and key1='inv_contract'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Contract Type defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtContract.Text = StrCodeReturn
                End If
            Case 12 'Shipment Mode
                StrCodeReturn = ShowList(1, 2000, Trim(txtShipmentMode.Text), " key2 ", " Descr ", " lists", " and key1='shipment_mode'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Mode Of Shipment defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtShipmentMode.Text = StrCodeReturn
                End If
            Case 13 'Dispatch Mode
                StrCodeReturn = ShowList(1, 2000, Trim(txtDispatchMode.Text), " key2 ", " Descr ", " lists", " and key1='dispatchmode'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Dispatch Mode defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtDispatchMode.Text = StrCodeReturn
                End If
            'Samiksha shipaddresscode changes
            Case 14 'Commodity
                StrCodeReturn = ShowList(1, 2000, Trim(txtCommodity.Text), " key2 ", " Descr ", " lists", " and key1='Types of commodities'")
                If StrCodeReturn = "-1" Then
                    MsgBox("No Commodity Type defined", MsgBoxStyle.Information, "empower")
                    Exit Sub
                Else
                    Me.txtCommodity.Text = StrCodeReturn
                End If

        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ShowValuestoString(ByRef pstrdispValues As String, ByRef pstrMode As String) As String
        Dim strArr() As String
        Dim intLoopCounter As Short
        mstrMode = pstrMode
        If pstrMode = "MODE_VIEW" Then
            For intLoopCounter = 3 To 13
                cmdHelp(intLoopCounter).Enabled = False
            Next
            DTPExchangeDate.Enabled = False : txtDeliveryTerms.Enabled = False : txtDeliveryTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtPaymentTerms.Enabled = False : TxtPaymentTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtInvoice_Desc_Buyer.Enabled = False : txtInvoice_Desc_Buyer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtInvoice_Decs_AEPC.Enabled = False : txtInvoice_Decs_AEPC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtFreight.Enabled = False
            txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtCurrency.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtOrigin_Status.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtCtryFinalDest.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtPreCarriage.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtPlaceOfReceipt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtPortOfLoading.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtPortOfDischarge.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtContract.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtFinalDest.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtShipmentMode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtDispatchMode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.txtotherref.Enabled = False : Me.txtotherref.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.txtbuyer.Enabled = False : Me.txtbuyer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtExchangeValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtExchangeValue.Enabled = False
            Txtadvlicno.Enabled = False : Txtadvlicno.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtPalLen.Enabled = False
            TxtPalLen.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtPalwid.Enabled = False
            TxtPalwid.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtPalhei.Enabled = False
            TxtPalhei.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtNoofPal.Enabled = False
            TxtNoofPal.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtARENo.Enabled = False : TxtARENo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtNetWei.Enabled = False
            TxtNetWei.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            TxtGrsWei.Enabled = False
            TxtGrsWei.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtvolweight.Enabled = False
            txtvolweight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            FramType.Enabled = False
            FramDrawback.Enabled = False
            TXTHSCODE.Enabled = False
            TXTHSCODE.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            'Samiksha commodity changes
            txtCommodity.Enabled = False
            txtCommodity.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

        Else
            For intLoopCounter = 3 To 13
                cmdHelp(intLoopCounter).Enabled = True
            Next
            DTPExchangeDate.Enabled = True : txtDeliveryTerms.Enabled = True : txtDeliveryTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtPaymentTerms.Enabled = True : TxtPaymentTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtInvoice_Desc_Buyer.Enabled = True : txtInvoice_Desc_Buyer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtInvoice_Decs_AEPC.Enabled = True : txtInvoice_Decs_AEPC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtFreight.Enabled = True
            txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtCurrency.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtOrigin_Status.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtCtryFinalDest.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtPreCarriage.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtPlaceOfReceipt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtPortOfLoading.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtPortOfDischarge.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtContract.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtFinalDest.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtShipmentMode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtDispatchMode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : Me.txtotherref.Enabled = True : Me.txtotherref.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Me.txtbuyer.Enabled = True : Me.txtbuyer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtExchangeValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtExchangeValue.Enabled = True
            If Trim(txtCurrency.Text) <> "" Then
                txtExchangeValue.Text = CStr(ShowExchangeRate(txtCurrency.Text))
                If UCase(Trim(txtCurrency.Text)) = UCase(gstrCURRENCYCODE) Then txtExchangeValue.Enabled = False
            Else
                txtExchangeValue.Text = "1.00"
            End If
            Txtadvlicno.Enabled = True : Txtadvlicno.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtPalLen.Enabled = True
            TxtPalLen.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtPalwid.Enabled = True
            TxtPalwid.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtPalhei.Enabled = True
            TxtPalhei.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtNoofPal.Enabled = True
            TxtNoofPal.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtARENo.Enabled = True : TxtARENo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtNetWei.Enabled = True
            TxtNetWei.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            TxtGrsWei.Enabled = True
            TxtGrsWei.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtvolweight.Enabled = True
            txtvolweight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            FramType.Enabled = True
            FramDrawback.Enabled = True
            TXTHSCODE.Enabled = True
            TXTHSCODE.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            'Samiksha commodity changes
            txtCommodity.Enabled = True
            txtCommodity.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        End If
        Select Case pstrMode
            Case "MODE_ADD"
                If Len(Trim(pstrdispValues)) = 0 Then
                    txtCurrency.Text = mstrCurrencyID : txtOrigin_Status.Text = ""
                    txtCtryFinalDest.Text = "" : txtPreCarriage.Text = ""
                    txtPlaceOfReceipt.Text = "" : txtPortOfLoading.Text = ""
                    txtPortOfLoading.Text = "" : txtPortOfDischarge.Text = ""
                    txtContract.Text = "" : txtFinalDest.Text = ""
                    txtShipmentMode.Text = "" : txtDispatchMode.Text = ""
                    txtDeliveryTerms.Text = "" : TxtPaymentTerms.Text = mstrCreditTermId
                    txtInvoice_Desc_Buyer.Text = ""
                    txtExchangeValue.Text = CStr(ShowExchangeRate(mstrCurrencyID))
                    txtFreight.Text = "0.00"
                    Me.txtotherref.Text = "" : Me.txtbuyer.Text = ""
                    If UCase(Trim(txtCurrency.Text)) = UCase(gstrCURRENCYCODE) Then
                        txtExchangeValue.Enabled = False
                        txtExchangeValue.Text = "1.00"
                    Else
                        txtExchangeValue.Enabled = True
                    End If
                    Txtadvlicno.Text = ""
                    TxtPalLen.Text = ""
                    TxtPalwid.Text = ""
                    TxtPalhei.Text = ""
                    TxtNoofPal.Text = ""
                    TxtARENo.Text = ""
                    TxtNetWei.Text = ""
                    TxtGrsWei.Text = ""
                    txtvolweight.Text = ""
                    Optpal.Checked = True
                    'optUnderAir.Checked = True
                    TXTHSCODE.Text = ""
                    'Samiksha commodity changes
                    txtCommodity.Text = ""
                Else
                    strArr = Split(pstrdispValues, "§")
                    txtCurrency.Text = strArr(0) : txtOrigin_Status.Text = strArr(1)
                    txtCtryFinalDest.Text = strArr(2) : txtPreCarriage.Text = strArr(3)
                    txtPlaceOfReceipt.Text = strArr(4) : txtPortOfLoading.Text = strArr(5)
                    txtPortOfDischarge.Text = strArr(6)
                    txtContract.Text = strArr(7) : txtFinalDest.Text = strArr(8)
                    txtShipmentMode.Text = strArr(9) : txtDispatchMode.Text = strArr(10)
                    txtDeliveryTerms.Text = strArr(11) : TxtPaymentTerms.Text = strArr(12)
                    txtInvoice_Desc_Buyer.Text = strArr(13)
                    txtInvoice_Decs_AEPC.Text = strArr(14)
                    txtExchangeValue.Text = strArr(15)
                    txtFreight.Text = strArr(16)
                    DTPExchangeDate.Value = VB6.Format(strArr(17), "MMM/yyyy")
                    Me.txtotherref.Text = strArr(18)
                    Me.txtbuyer.Text = strArr(19)
                    If UCase(Trim(txtCurrency.Text)) = UCase(gstrCURRENCYCODE) Then
                        txtExchangeValue.Enabled = False
                        txtExchangeValue.Text = "1.00"
                    Else
                        txtExchangeValue.Enabled = True
                    End If
                    Txtadvlicno.Text = strArr(20)
                    TxtPalLen.Text = strArr(21)
                    TxtPalwid.Text = strArr(22)
                    TxtPalhei.Text = strArr(23)
                    TxtNoofPal.Text = strArr(24)
                    TxtARENo.Text = strArr(25)
                    TxtNetWei.Text = strArr(26)
                    TxtGrsWei.Text = strArr(27)
                    If strArr(28) = "P" Then
                        Optpal.Checked = True
                    Else
                        OptCon.Checked = True
                    End If
                    txtvolweight.Text = strArr(29)
                    If strArr(30) = "A" Then
                        CHKUnderAir.Checked = True
                        CHKUnderBRF.Checked = False
                    ElseIf strArr(30) = "B" Then
                        CHKUnderBRF.Checked = True
                        CHKUnderAir.Checked = False
                    Else
                        CHKUnderAir.Checked = False
                        CHKUnderBRF.Checked = False
                    End If
                    'Samiksha commodity changes
                    txtCommodity.Text = strArr(34)
                End If
            Case "MODE_VIEW", "MODE_EDIT"
                strArr = Split(pstrdispValues, "§")
                txtCurrency.Text = strArr(0) : txtOrigin_Status.Text = strArr(1)
                txtCtryFinalDest.Text = strArr(2) : txtPreCarriage.Text = strArr(3)
                txtPlaceOfReceipt.Text = strArr(4) : txtPortOfLoading.Text = strArr(5)
                txtPortOfDischarge.Text = strArr(6)
                txtContract.Text = strArr(7) : txtFinalDest.Text = strArr(8)
                txtShipmentMode.Text = strArr(9) : txtDispatchMode.Text = strArr(10)
                txtDeliveryTerms.Text = strArr(11) : TxtPaymentTerms.Text = strArr(12)
                txtInvoice_Desc_Buyer.Text = strArr(13)
                txtInvoice_Decs_AEPC.Text = strArr(14)
                txtExchangeValue.Text = strArr(15)
                txtFreight.Text = strArr(16)
                DTPExchangeDate.Value = VB6.Format(VB6.Format(strArr(17), gstrDateFormat), "MMM/yyyy")
                Me.txtotherref.Text = strArr(18)
                Me.txtbuyer.Text = strArr(19)
                If UBound(strArr) >= 20 Then Txtadvlicno.Text = strArr(20)
                If UBound(strArr) >= 21 Then TxtPalLen.Text = strArr(21)
                If UBound(strArr) >= 22 Then TxtPalwid.Text = strArr(22)
                If UBound(strArr) >= 23 Then TxtPalhei.Text = strArr(23)
                If UBound(strArr) >= 24 Then TxtNoofPal.Text = strArr(24)
                If UBound(strArr) >= 25 Then TxtARENo.Text = strArr(25)
                If UBound(strArr) >= 26 Then TxtNetWei.Text = strArr(26)
                If UBound(strArr) >= 27 Then TxtGrsWei.Text = strArr(27)
                If UBound(strArr) >= 28 Then
                    If strArr(28) = "P" Then
                        Optpal.Checked = True
                    Else
                        OptCon.Checked = True
                    End If
                End If
                If UBound(strArr) >= 29 Then txtvolweight.Text = strArr(29)

                If UBound(strArr) >= 30 Then
                    If strArr(30) = "A" Then
                        CHKUnderAir.Checked = True
                    ElseIf strArr(30) = "B" Then
                        CHKUnderBRF.Checked = True
                    Else
                        CHKUnderBRF.Checked = False
                        CHKUnderAir.Checked = False
                    End If
                End If

                If UBound(strArr) >= 31 Then TXTHSCODE.Text = strArr(31)
                'Samiksha commodity code changes
                If UBound(strArr) >= 32 Then txtCommodity.Text = strArr(32)

        End Select
        Me.ShowDialog()
    End Function
    Public Function AddValuestoString() As String

        Dim Strunder As String
        If CHKUnderAir.Checked = True Then
            Strunder = "A"
        ElseIf CHKUnderBRF.Checked = True Then
            Strunder = "B"
        Else
            Strunder = ""
        End If
        strValues = ""
        strValues = Trim(txtCurrency.Text) & "§" & Trim(txtOrigin_Status.Text) & "§"
        strValues = strValues & Trim(txtCtryFinalDest.Text) & "§" & Trim(txtPreCarriage.Text) & "§"
        strValues = strValues & Trim(txtPlaceOfReceipt.Text) & "§"
        strValues = strValues & Trim(txtPortOfLoading.Text) & "§" & Trim(txtPortOfDischarge.Text) & "§"
        strValues = strValues & Trim(txtContract.Text) & "§" & Trim(txtFinalDest.Text) & "§"
        strValues = strValues & Trim(txtShipmentMode.Text) & "§" & Trim(txtDispatchMode.Text) & "§"
        strValues = strValues & Trim(txtDeliveryTerms.Text) & "§" & Trim(TxtPaymentTerms.Text) & "§"
        strValues = strValues & txtInvoice_Desc_Buyer.Text & "§" & txtInvoice_Decs_AEPC.Text & "§"
        strValues = strValues & CStr(txtExchangeValue.Text) & "§" & CStr(txtFreight.Text)
        strValues = strValues & "§" & CStr(Me.DTPExchangeDate.Value) & "§" & Trim(Me.txtotherref.Text) & "§" & Trim(Me.txtbuyer.Text)
        strValues = strValues & "§" & Trim(Txtadvlicno.Text) & "§" & Trim(TxtPalLen.Text) & "§" & Trim(TxtPalwid.Text)
        strValues = strValues & "§" & Trim(TxtPalhei.Text) & "§" & Val(TxtNoofPal.Text) & "§" & Trim(TxtARENo.Text)
        strValues = strValues & "§" & Val(TxtNetWei.Text) & "§" & Val(TxtGrsWei.Text) & "§" & IIf(Optpal.Checked, "P", "C")
        strValues = strValues & "§" & Val(txtvolweight.Text) & "§" & Strunder & "§" & Trim(TXTHSCODE.Text)
        'Samiksha commodity changes
        strValues = strValues & "§" & Val(txtCommodity.Text) & "§" & Strunder & "§" & Trim(txtCommodity.Text)

        AddValuestoString = strValues
    End Function
    Private Function ValidatebeforeSave() As Boolean
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        ValidatebeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        If (Len(Trim(txtCurrency.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Currency Code."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(3)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtOrigin_Status.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Ctry of Origin."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(4)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtCtryFinalDest.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Ctry of final Dest."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(5)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtPreCarriage.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ".Pre Carriage "
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(6)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtPlaceOfReceipt.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Place Of Receipt."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(7)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtPortOfLoading.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Port of Loading."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(8)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtPortOfDischarge.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Port of Discharge."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(9)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtContract.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Contract."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(11)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtFinalDest.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Final Destination."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(10)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtShipmentMode.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Shipment Mode."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(12)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtDispatchMode.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Dispatch Mode."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = cmdHelp(13)
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(txtInvoice_Decs_AEPC.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". EPGC No."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = txtInvoice_Decs_AEPC
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(Txtadvlicno.Text)) = 0) And mstrInvSubType = "EXPORTS" Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Advance Licence No."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Txtadvlicno
            End If
            ValidatebeforeSave = False
        End If
        If Optpal.Checked = True Then
            If (Len(Trim(TxtPalLen.Text)) = 0) Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Pallet Length."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = TxtPalLen
                End If
                ValidatebeforeSave = False
            End If
            If (Len(Trim(TxtPalwid.Text)) = 0) Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Pallet Width."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = TxtPalwid
                End If
                ValidatebeforeSave = False
            End If
            If (Len(Trim(TxtPalhei.Text)) = 0) Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Pallet Height."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = TxtPalhei
                End If
                ValidatebeforeSave = False
            End If
            If (Len(Trim(TxtNoofPal.Text)) = 0) Then
                lstrControls = lstrControls & vbCrLf & lNo & ". No Of Pallets."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = TxtNoofPal
                End If
                ValidatebeforeSave = False
            End If
        End If
        If (Len(Trim(TxtARENo.Text)) = 0) And mstrInvSubType = "EXPORTS" Then
            lstrControls = lstrControls & vbCrLf & lNo & ". ARE No."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = TxtARENo
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(TxtNetWei.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Net Weight."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = TxtNetWei
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Trim(TxtGrsWei.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Gross Weight."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = TxtGrsWei
            End If
            ValidatebeforeSave = False
        End If

        If (Len(Trim(TXTHSCODE.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". HS CODE."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = TXTHSCODE
            End If
            ValidatebeforeSave = False
        End If
        'Samiksha commodity changes

        If (Len(Trim(txtCommodity.Text)) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Commodty."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = txtCommodity
            End If
            ValidatebeforeSave = False
        End If

        If Not ValidatebeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            If lctrFocus.Enabled = True Then lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Function
    End Function
    Private Sub txtbuyer_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtbuyer.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                CmdOk.Focus()
            Case 39, 34, 96
                KeyAscii = 0
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
    Private Sub TxtCurrency_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCurrency.TextChanged
        If Trim(mstrMode) <> "MODE_VIEW" Then
            If UCase(Trim(txtCurrency.Text)) = UCase(gstrCURRENCYCODE) Then
                txtExchangeValue.Enabled = False
                txtExchangeValue.Text = "1.00"
            Else
                txtExchangeValue.Enabled = True
            End If
        End If
    End Sub
    Private Sub txtDeliveryTerms_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDeliveryTerms.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                TxtPaymentTerms.Focus()
            Case 39, 34, 96
                KeyAscii = 0
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
    Private Sub txtExchangeValue_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtExchangeValue.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 46 Or KeyAscii = 8 Then
            If InStr(1, Trim(txtExchangeValue.Text), ".") <> 0 And KeyAscii = 46 Then
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtExchangeValue_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtExchangeValue.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If Not IsNumeric(Trim(txtExchangeValue.Text)) Then
            txtExchangeRate.Text = "1.00"
        End If
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtInvoice_Decs_AEPC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice_Decs_AEPC.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtotherref.Focus()
            Case 39, 34, 96
                KeyAscii = 0
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
    Private Sub txtInvoice_Desc_Buyer_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice_Desc_Buyer.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtInvoice_Decs_AEPC.Focus()
            Case 39, 34, 96
                KeyAscii = 0
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
    Private Sub txtotherref_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtotherref.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtbuyer.Focus()
            Case 39, 34, 96
                KeyAscii = 0
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
    Private Sub txtPaymentTerms_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtPaymentTerms.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtInvoice_Desc_Buyer.Focus()
            Case 39, 34, 96
                KeyAscii = 0
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
    Public Property SetDocumentDate() As String
        Get
            SetDocumentDate = mstrDocumentDate
        End Get
        Set(ByVal Value As String)
            mstrDocumentDate = Value
        End Set
    End Property
    Public Property SetCurrencyID() As String
        Get
            SetCurrencyID = mstrCurrencyID
        End Get
        Set(ByVal Value As String)
            mstrCurrencyID = Value
        End Set
    End Property
    Public Property SetCreditTermID() As String
        Get
            SetCreditTermID = mstrCreditTermId
        End Get
        Set(ByVal Value As String)
            mstrCreditTermId = Value
        End Set
    End Property
    
    Private Function ShowExchangeRate(ByVal pstrCurrencyID As String) As Double
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Created By     -   Tapan Jain
        'Description    -   Get Data from BackEnd
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ADODB.Recordset
        Dim strTempDocDate As String
        On Error GoTo ErrHandler
        If Trim(mstrDocumentDate) <> "" Then
            strTempDocDate = VB6.Format(mstrDocumentDate, gstrDateFormat)
        Else
            strTempDocDate = GetServerDate()
        End If
        StrSQLQuery = "SET DATEFORMAT 'mdy'" & vbCrLf & "SELECT CExch_MultiFactor From Gen_CurExchMaster Where unit_code='" & gstrUNITID & "' and CExch_CurrencyTo='" & Trim(pstrCurrencyID) & "' AND CExch_InOut=1 AND '" & getDateForDB(strTempDocDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo "
        'If GetDataFromTable.State = 1 Then GetDataFromTable.Close()
        GetDataFromTable = New ADODB.Recordset
        GetDataFromTable.Open(StrSQLQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Not GetDataFromTable.EOF Or Not GetDataFromTable.BOF Then
            ShowExchangeRate = GetDataFromTable.Fields("CExch_MultiFactor").Value
        Else
            txtCurrency.Text = ""
            ShowExchangeRate = 1.0#
        End If
        GetDataFromTable.Close()
        GetDataFromTable = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        ShowExchangeRate = 1.0#
    End Function

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TxtARENo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtARENo.TextChanged

    End Sub

    Private Sub CHKUnderAir_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHKUnderAir.CheckStateChanged
        If CHKUnderAir.CheckState = System.Windows.Forms.CheckState.Checked Then
            CHKUnderBRF.Checked = False
            CHKUnderAir.Checked = True
        End If
    End Sub

    Private Sub CHKUnderBRF_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHKUnderBRF.CheckStateChanged
        If CHKUnderBRF.CheckState = System.Windows.Forms.CheckState.Checked Then
            CHKUnderBRF.Checked = True
            CHKUnderAir.Checked = False
        End If

    End Sub

    Private Sub CHKUnderAir_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKUnderAir.CheckedChanged

    End Sub


End Class