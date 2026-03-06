Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0031
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'Copyright          :   MIND Ltd.
	'Form Name          :   frmMKTTRN0031.frm
	'Created By         :   Arshad Ali
	'Created on         :   15/04/2004
	'Modified Date      :
    'Description        :   This form is used to show items to select from.

    'Modified By Nitin Mehta on 13 May 2011
    'Modified to support MultiUnit functionality
	'---------------------------------------------------------------------------
	Dim mstrItemText As String
	Dim mstrInvType As String
    Dim mstrInvSubType As String
    Dim strPar1 As String
    Dim strPar2, strPar3, strPar4, strPar5, strPar6 As String
    Dim Bool_show As Boolean = False
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrHandler
		Me.Close()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		'Created By   -   Arshad Ali
		'retrieve item code of all selected items
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		On Error GoTo ErrHandler
		mstrItemText = ""
		Dim intSubItem As Short
		With spItems
			For intSubItem = 1 To .maxRows
				.Row = intSubItem
				.Col = 1
				If CBool(.Value) = True Then
					.Col = 2
					mstrItemText = mstrItemText & "'" & Trim(.Text) & "',"
				End If
			Next intSubItem
		End With
		If Len(mstrItemText) = 0 Then
			Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
			Me.spItems.Focus()
			Exit Sub
		End If
		Me.Close()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	
    Private Sub frmMKTTRN0031_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Change cursor to arrow when over headers
        SpItems.CursorType = FPSpreadADO.CursorTypeConstants.CursorTypeColHeader
        SpItems.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        SpItems.Focus()
    End Sub
	
    Sub AddColumnsInSpread()
        With SpItems
            .MaxRows = 0
            .MaxCols = 4
            .Row = 0
            .Col = 1 : .Text = "Mark" : .set_ColWidth(1, 4)
            .Col = 2 : .Text = "Item Code" : .set_ColWidth(2, 12)
            .Col = 3 : .Text = "Description" : .set_ColWidth(3, 20)
            .Col = 4 : .Text = "Tariff Code" : .set_ColWidth(4, 7)
        End With
    End Sub

	Public Function SelectDatafromItem_Mst(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrstockLocation As String, Optional ByRef pstrAccountCode As String = "", Optional ByRef pstrItemNotin As String = "", Optional ByRef intAlreadyItem As Short = 0) As Object
		On Error GoTo ErrHandler
		Dim strItembal As String
		Dim rsItembal As ClsResultSetDB
		Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short

        Call AddColumnsInSpread()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optDescription.Checked = True
        mstrItemText = ""

        Select Case pstrInvType
            Case "NORMAL INVOICE"
                Select Case pstrInvSubtype
                    Case "TRADING GOODS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp ='T'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "ASSETS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp ='P'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "TOOLS & DIES"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp in('P','A')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "RAW MATERIAL"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp IN('C','R','B','M')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "SCRAP"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE and a.Item_Code in (Select Item_Code  from ItemBal_Mst Where Location_Code ='" & pstrstockLocation & "' and cur_Bal > 0 AND UNIT_CODE='" & gstrUNITID & "')"
                        strItembal = strItembal & " and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "COMPONENTS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp ='C'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "FINISHED GOODS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp = 'F'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "ALL"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.UNIT_CODE=b.UNIT_CODE "
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                End Select
            Case "SAMPLE INVOICE"
                Select Case pstrInvSubtype
                    Case "FINISHED GOODS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code AND a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp = 'F'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "RAW MATERIAL"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code AND a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp ='R'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "COMPONENTS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code AND a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp ='C'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                End Select
            Case "TRANSFER INVOICE"
                Select Case pstrInvSubtype
                    Case "ASSETS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code AND a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp ='P'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "FINISHED GOODS"
                        strItembal = "SELECT Distinct a.Item_Code,c.Cust_drgNo,c.Drg_Desc, a.Tariff_code FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE=c.UNIT_CODE and a.Item_Main_Grp = 'F' and a.Item_Code = c.ITem_Code"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & pstrAccountCode & "' AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "INPUTS"
                        strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.Item_Code=b.Item_Code AND a.UNIT_CODE=b.UNIT_CODE and a.Item_Main_Grp in('R','C','M','N','S','B','A')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                End Select
            Case "REJECTION"
                strItembal = "SELECT Distinct(a.Item_Code),a.description,c.Tariff_code FROM vend_item a ,Itembal_Mst b,Item_Mst c"
                strItembal = strItembal & " where a.Item_Code=b.Item_Code and a.Item_code = c.Item_code AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE=c.UNIT_CODE  and a.Account_code ='" & pstrAccountCode & "' AND a.UNIT_CODE='" & gstrUNITID & "' "
                strItembal = strItembal & " and cur_bal >0 "
                strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                If Len(Trim(pstrItemNotin)) > 0 Then
                    strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                End If
        End Select
        rsItembal = New ClsResultSetDB
        If Len(Trim(strItembal)) <= 0 Then Exit Function
        rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rsItembal.MoveFirst() 'move to first record
            If (UCase(pstrInvType) = "TRANSFER INVOICE") And UCase(pstrInvSubtype) = "FINISHED GOODS" Then
                For intCount = 1 To intRecordCount
                    With SpItems
                        .MaxRows = .MaxRows + 1
                        .Row = intCount
                        .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
                        .Col = 2

                        .Text = rsItembal.GetValue("Item_code") : .Lock = True
                        .Col = 3

                        .Text = rsItembal.GetValue("Drg_Desc") : .Lock = True
                        .Col = 4

                        .Text = rsItembal.GetValue("Tariff_Code") : .Lock = True
                    End With
                    rsItembal.MoveNext() 'move to next record
                Next intCount
            Else
                For intCount = 1 To intRecordCount
                    With SpItems
                        .MaxRows = .MaxRows + 1
                        .Row = intCount
                        .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
                        .Col = 2

                        .Text = rsItembal.GetValue("Item_code") : .Lock = True
                        .Col = 3

                        .Text = rsItembal.GetValue("Description") : .Lock = True
                        .Col = 4

                        .Text = rsItembal.GetValue("Tariff_Code") : .Lock = True
                    End With
                    rsItembal.MoveNext() 'move to next record
                Next intCount
            End If
            rsItembal.ResultSetClose()

            rsItembal = Nothing
        Else
            If (UCase(pstrInvType) = "TRANSFER INVOICE") And UCase(pstrInvSubtype) = "FINISHED GOODS" Then
                MsgBox("No items details defined  for above Invoice combination,Please Check Following :" & vbCrLf & "1. Item should be Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3.Item is not defined in Customer ITem Master.", MsgBoxStyle.Information, "eMPro")
            Else
                MsgBox("No items details defined  for above Invoice combination,Please Check Following :" & vbCrLf & "1. Item should be Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & ".", MsgBoxStyle.Information, "eMPro")
            End If
            Exit Function
        End If
        Me.ShowDialog()

        SelectDatafromItem_Mst = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

    Private Sub SpItems_KeyUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpItems.KeyUpEvent
        With SpItems
            If eventArgs.keyCode = 13 Or eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
                .Col = 1
                .Value = IIf(CBool(.Value), False, True)
            End If
        End With
    End Sub

    Private Sub SpItems_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpItems.LeaveCell
        With SpItems
            .Row = -1
            .Col = -1
            .BackColor = System.Drawing.Color.White
            .ForeColor = System.Drawing.Color.Black
            .Col = -1
            .Row = IIf(eventArgs.newRow <= 0, 1, eventArgs.newRow)
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000D)
            .ForeColor = System.Drawing.Color.White
        End With
    End Sub

    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
    End Sub

    Sub SearchItem()
        On Error GoTo ErrHandler
        Dim intCount As Short
        With SpItems
            .Row = -1
            .Col = -1
            .BackColor = System.Drawing.Color.White
            .ForeColor = System.Drawing.Color.Black
            If optItemCode.Checked Then
                .Col = 2
            End If
            If optDescription.Checked Then
                .Col = 3
            End If
            If optTariff.Checked Then
                .Col = 4
            End If
            For intCount = 1 To .MaxRows
                .Row = intCount
                If UCase(Mid(.Text, 1, Len(txtsearch.Text))) = UCase(txtsearch.Text) Then
                    .TopRow = .Row
                    .Col = -1
                    .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000D)
                    .ForeColor = System.Drawing.Color.White
                    Exit Sub
                End If
            Next
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtSearch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtsearch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtsearch_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtsearch.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        With SpItems
            If KeyCode = 13 And Len(Trim(txtsearch.Text)) > 0 Then
                .Col = 1
                .Value = IIf(CBool(.Value), False, True)
            End If
        End With
    End Sub

    Public Function SelectDataFromCustOrd_Dtl(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrSubType As String, ByRef pstrInvType As String, ByRef pstrstockLocation As String, Optional ByRef pstrCondition As String = "", Optional ByRef intAlreadyItem As Short = 0) As String
        '***********************************
        'To Get Data From Cust_Ord_Dtl
        '***********************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim Validyrmon As String
        Dim effectyrmon As String
        Dim validMon As String
        Dim effectMon As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        Dim strDate As String

        Call AddColumnsInSpread()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optDescription.Checked = True
        mstrItemText = ""

        strDate = setDateFormat(GetServerDate()) 'VB6.Format(GetServerDate, gstrDateFormat)

        strSelectSql = "Select effectMon=convert(char(2),month(effect_date)),effectYr=convert(char(4),Year(effect_date)),"
        strSelectSql = strSelectSql & " validMon=convert(char(2),month(Valid_date)),validYr=convert(char(4),year(Valid_date))"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr where "
        strSelectSql = strSelectSql & " Account_Code='" & Trim(pstrCustno) & "' and Cust_Ref='" & Trim(pstrRefNo) & "'"
        strSelectSql = strSelectSql & " and Amendment_No='" & Trim(pstrAmmNo) & "' and Active_Flag = 'A' AND UNIT_CODE='" & gstrUNITID & "'"
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdHdr.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustOrdHdr.GetNoRows > 0 Then
            validMon = CStr(Month(GetServerDate))
            If CDbl(validMon) < 10 Then
                validMon = "0" & validMon
            End If

            Validyrmon = Year(GetServerDate) & validMon

            effectMon = rsCustOrdHdr.GetValue("EffectMon")
            If CDbl(effectMon) < 10 Then
                effectMon = "0" & effectMon
            End If

            effectyrmon = rsCustOrdHdr.GetValue("effectYr") & effectMon
            mstrInvType = pstrInvType : mstrInvSubType = pstrSubType

            Select Case UCase(pstrInvType)
                Case "NORMAL INVOICE", "EXPORT INVOICE", "SERVICE INVOICE"
                    Select Case UCase(pstrSubType)
                        Case "FINISHED GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F'", pstrCondition)
                        Case "COMPONENTS"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'C'", pstrCondition)
                        Case "RAW MATERIAL"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'R','S','B','M'", pstrCondition)
                        Case "ASSETS"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P'", pstrCondition)
                        Case "TRADING GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'T'", pstrCondition)
                        Case "ALL"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'M','N','F','C','S','B','R','P','A'", pstrCondition)
                        Case "TOOLS & DIES"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P','A'", pstrCondition)
                        Case "EXPORTS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition)
                        Case "SERVICE"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'F','S'", pstrCondition)
                    End Select
                Case "JOBWORK INVOICE"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F'", pstrCondition)
            End Select
        Else
            rsCustOrdHdr.ResultSetClose()

            rsCustOrdHdr = Nothing
            strSelectSql = "Select effect_date,"
            strSelectSql = strSelectSql & " Valid_date,"
            strSelectSql = strSelectSql & " from Cust_Ord_hdr where "
            strSelectSql = strSelectSql & " Account_Code='" & Trim(pstrCustno) & "' and Cust_Ref='" & Trim(pstrRefNo) & "'"
            strSelectSql = strSelectSql & " and Amendment_No='" & Trim(pstrAmmNo) & "' and Active_flag ='A' AND UNIT_CODE='" & gstrUNITID & "'"
            rsCustOrdHdr = New ClsResultSetDB
            rsCustOrdHdr.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustOrdHdr.GetNoRows > 0 Then

                Validyrmon = rsCustOrdHdr.GetValue("valid_date")

                effectyrmon = rsCustOrdHdr.GetValue("Effect_date")
            End If
            rsCustOrdHdr.ResultSetClose()
            rsCustOrdHdr = Nothing
            Select Case pstrSubType
                Case "COMPONENTS"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'C'", pstrCondition)
                Case "TRADING GOODS"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'T'", pstrCondition)
                Case "ASSETS"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'P'", pstrCondition)
                Case "TOOLS & DIES"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'A','P'", pstrCondition)
                Case "RAW MATERIAL"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','S','B','M'", pstrCondition)
                Case "SCRAP"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','C'", pstrCondition)
            End Select
        End If
        rsCustOrdDtl = New ClsResultSetDB

        rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsCustOrdDtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rsCustOrdDtl.MoveFirst() 'move to first record
            For intCount = 1 To intRecordCount
                With SpItems
                    .MaxRows = .MaxRows + 1
                    .Row = intCount
                    .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
                    .Col = 2

                    .Text = rsCustOrdDtl.GetValue("Cust_Drgno") : .Lock = True
                    .Col = 3

                    .Text = rsCustOrdDtl.GetValue("Cust_Drg_Desc") : .Lock = True
                    .Col = 4

                    .Text = rsCustOrdDtl.GetValue("Tariff_Code") : .Lock = True
                End With
                rsCustOrdDtl.MoveNext() 'move to next record
            Next intCount
            rsCustOrdDtl.ResultSetClose()

            rsCustOrdDtl = Nothing
        Else
            MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check Marketing Schedule in Case of Finished\Trading Goods in SO.", MsgBoxStyle.Information, "eMPro")
            Exit Function
        End If
        Me.ShowDialog()
        SelectDataFromCustOrd_Dtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

    Public Function makeSelectSql(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "") As String
        Dim strSelectSql As String
        strSelectSql = "Select b.Item_Code, d.item_code as Cust_DrgNo, d.description as Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = c.UNIT_CODE AND a.UNIT_CODE = d.UNIT_CODE and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code  and c.Cust_drgNo=b.Cust_drgNo  and b.ITem_code = d.Item_code  and a.Account_Code='" & Trim(pstrCustno) & "' AND a.UNIT_CODE='" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code AND a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " UNION "
        strSelectSql = strSelectSql & "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = c.UNIT_CODE AND a.UNIT_CODE = d.UNIT_CODE AND a.UNIT_CODE = '" & gstrUNITID & "' AND a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code  and c.Cust_drgNo=b.Cust_drgNo  and b.ITem_code =d.ITem_code  and b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and  datepart(mm,b.trans_date) = " & Month(ConvertToDate(strDate))
        strSelectSql = strSelectSql & " and datepart(dd,b.trans_date) <= " & ConvertToDate(strDate).Day & " and  datepart(yyyy,b.trans_date) = " & Year(ConvertToDate(strDate))   'Mid(strDate, 7, 4)
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and a.unit_code = b.unit_code  AND a.UNIT_CODE='" & gstrUNITID & "' and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        makeSelectSql = strSelectSql
    End Function

    Public Function AddDataFromGrinDtl(ByRef pstrVend As String, ByRef dblGrnNo As Double, ByRef pstrstockLocation As String, Optional ByRef intAlreadyItem As Short = 0, Optional ByRef pstrCondition As String = "") As String
        Dim rsGrnDtl As ClsResultSetDB
        Dim strSQL As String
        Dim StrItemCode As String
        Dim strItemNot As String
        Dim arrRejAcpt(,) As Object
        Dim intLoopCounter As Short
        Dim intArrLoopCount As Short
        Dim intmaxLoop As Short
        Dim intUbound As Short
        Dim intCount As Short
        mstrInvType = "REJECTION" : mstrInvSubType = "REJECTION"

        On Error GoTo ErrHandler

        Call AddColumnsInSpread()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optDescription.Checked = True
        mstrItemText = ""


        rsGrnDtl = New ClsResultSetDB
        strSQL = "select a.Doc_No,a.Item_code,a.Rejected_Quantity,Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
        strSQL = strSQL & " Inspected_Quantity = isnull(a.Inspected_Quantity,0), RGP_Quantity = isnull(a.RGP_Quantity,0)  from grn_Dtl a,"
        strSQL = strSQL & " grn_hdr b Where "
        strSQL = strSQL & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
        strSQL = strSQL & "a.From_Location = b.From_Location AND a.UNIT_CODE=b.UNIT_CODE and a.From_Location ='01R1'"
        strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
        strSQL = strSQL & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 AND a.UNIT_CODE='" & gstrUNITID & "' "
        If Len(Trim(pstrCondition)) > 0 Then
            strSQL = strSQL & " and a.Item_code not in (" & pstrCondition & ")"
        End If
        rsGrnDtl.GetResult(strSQL)
        If rsGrnDtl.GetNoRows > 0 Then
            intmaxLoop = rsGrnDtl.GetNoRows : rsGrnDtl.MoveFirst() : ReDim arrRejAcpt(2, intmaxLoop - 1) : intUbound = intmaxLoop - 1
            '****To Fatch all Doc_No and Rejected Quantity in Array
            intUbound = intmaxLoop - 1
            For intLoopCounter = 1 To intmaxLoop
                arrRejAcpt(0, intLoopCounter - 1) = rsGrnDtl.GetValue("Item_Code")
                arrRejAcpt(1, intLoopCounter - 1) = rsGrnDtl.GetValue("Rejected_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
                rsGrnDtl.MoveNext()
            Next
            '****
            strItemNot = ""
            For intArrLoopCount = 0 To intUbound
                StrItemCode = arrRejAcpt(0, intArrLoopCount)
                If arrRejAcpt(1, intArrLoopCount) <= 0 Then
                    If Len(Trim(strItemNot)) > 0 Then
                        strItemNot = StrItemCode & ",'" & StrItemCode & "'"
                    Else
                        strItemNot = "'" & StrItemCode & "'"
                    End If
                End If
            Next
            If Len(Trim(strItemNot)) > 0 Then
                strSQL = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.Tariff_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                strSQL = strSQL & "a.Doc_type = b.Doc_type AND a.UNIT_CODE = b.UNIT_CODE and a.Doc_no = b.Doc_No "
                strSQL = strSQL & "and a.From_Location = b.From_Location "
                strSQL = strSQL & " and a.Item_Code = c.ITem_code AND a.UNIT_CODE = c.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' and b.From_Location ='01R1'"
                strSQL = strSQL & " and a.Item_code Not in (" & strItemNot & ")"
                strSQL = strSQL & " and c.Status = 'A' and Hold_Flag =0"
                strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                strSQL = strSQL & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 "
                strSQL = strSQL & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where UNIT_CODE='" & gstrUNITID & "' AND Location_Code = '"
                strSQL = strSQL & pstrstockLocation & "' and Cur_bal > 0)"
                If Len(Trim(pstrCondition)) > 0 Then
                    strSQL = strSQL & " and a.Item_code not in (" & pstrCondition & ")"
                End If
            Else
                strSQL = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.Tariff_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                strSQL = strSQL & "a.Doc_type = b.Doc_type AND a.UNIT_CODE = b.UNIT_CODE and a.Doc_no = b.Doc_No "
                strSQL = strSQL & "and a.From_Location = b.From_Location "
                strSQL = strSQL & " and a.Item_Code = c.ITem_code AND a.UNIT_CODE = c.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' and b.From_Location ='01R1'"
                strSQL = strSQL & " and c.Status = 'A' and Hold_Flag =0"
                strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                strSQL = strSQL & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 "
                strSQL = strSQL & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where UNIT_CODE='" & gstrUNITID & "' AND Location_Code = '"
                strSQL = strSQL & pstrstockLocation & "' and Cur_bal > 0)"
                If Len(Trim(pstrCondition)) > 0 Then
                    strSQL = strSQL & " and a.Item_code not in (" & pstrCondition & ")"
                End If
            End If
            rsGrnDtl.ResultSetClose()
            rsGrnDtl = New ClsResultSetDB
            rsGrnDtl.GetResult(strSQL)
            intmaxLoop = rsGrnDtl.GetNoRows 'assign record count to integer variable
            If intmaxLoop > 0 Then '          'if record found
                rsGrnDtl.MoveFirst() 'move to first record
                For intLoopCounter = 1 To intmaxLoop
                    With SpItems
                        .MaxRows = .MaxRows + 1
                        .Row = intCount
                        .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
                        .Col = 2

                        .Text = rsGrnDtl.GetValue("Item_code") : .Lock = True
                        .Col = 3

                        .Text = rsGrnDtl.GetValue("Description") : .Lock = True
                        .Col = 4

                        .Text = rsGrnDtl.GetValue("Tariff_Code") : .Lock = True
                    End With
                    rsGrnDtl.MoveNext() 'move to next record
                Next
            Else
                MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in Grin are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check supplimentry Grin for items in Grin(Selected) ", MsgBoxStyle.Information, "eMPro")
            End If
        End If
        rsGrnDtl.ResultSetClose()

        rsGrnDtl = Nothing
        Me.ShowDialog()
        AddDataFromGrinDtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
	
    Public Function MakeSelectSubQuery(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        Dim strSelectSql As String
        strSelectSql = "Select c.Item_Code, d.item_code as Cust_DrgNo,d.description as Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref AND a.UNIT_CODE = c.UNIT_CODE and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code AND c.UNIT_CODE = d.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref='" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' And c.Active_Flag = 'A' "
        strSelectSql = strSelectSql & " and c.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code AND a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        MakeSelectSubQuery = strSelectSql
    End Function
End Class