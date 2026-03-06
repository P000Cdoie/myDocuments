Option Strict Off
Option Explicit On
Friend Class frmMKTMST0002
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------
	'Copyright (c) - MIND
	'Name of module -  frmMKTMST0002.frm
	'Created By     -  Kapil
	'Created Date   -  30 - 04 - 2001
	'Description    -  Calender Master
	'Revised date   -  1.)29 - 04 - 2002
	'                  2.)26 - 03 - 2003 (Nitin Sood)
    'Modified by Sameer Srivastava on 2011-Apr-25
    '   Modified to support MultiUnit functionality
	'----------------------------------------------------
	
	Dim mintIndex As Short 'hold form count
	Dim mintDaysInMonth As Short 'hold days of the month
	Dim mstrStartFinYear As String 'to get the Current Financial Year
	Dim mintMonthIndex As Short 'to get Index Of Month i.e April=4,Jan=1 etc.
	Dim mstrYear As String 'to get Current Finncial Year for the Selected Month
	Dim mblnhasSave As Boolean
	Dim mintCheckPreMonth As Short 'restrict user so that he can't make changes to the previous month from current month
    Private Sub cbocalYear_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbocalYear.SelectedIndexChanged
        '-------------------------
        'Created By     -  Nitin Sood
        'Description    -  Display the Calender for the selected Year
        '-------------------------
        mstrStartFinYear = Mid(cbocalYear.Text, 1, 4)
        Call cmbCalMonth_SelectedIndexChanged(cmbCalMonth, New System.EventArgs())
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub cmbCalMonth_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCalMonth.SelectedIndexChanged
        '-------------------------
        'Created By     -  Kapil
        'Description    -  Display the Calender for the selected Month for Current Financial Year
        '-------------------------
        'On Error GoTo ErrHandler
        Call ClearColumnOfSpreadControl("Clear") 'Initially Clear all the Columns Of the Grid
        Call DisplayDaysOfMonth() 'Procedure Call to Display the Days Of the Month
        Call ClearColumnOfSpreadControl("View") 'Center Align the Text In the Spread and Make it Static text
        Call ClearColumnOfSpreadControl("BackColor") 'Set the Cell BackColor Of Sundays In the Spread
        Call CheckSundaysInDataBase() 'To check if Sunday is not Holiday in Company_Mst
        'then Set the Fore Color Of that Column Black
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub

    Private Sub cmbCalMonth_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbCalMonth.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '-------------------------
        'Created By     -  Kapil
        'Description    -  If Enter Key is Pressed then Display the Calender for the selected Month for Current Financial Year
        '-------------------------
        'On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Me.CmdGrpMkt.Focus()
                System.Windows.Forms.SendKeys.Send(vbTab)
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("Calender Master.htm")
    End Sub
    Private Sub frmMKTMST0002_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '-------------------------
        'Created By     -  Kapil
        'Description    -  Check the form in the MDIMenu when if it is activated i.e opened in MDIWindow
        '-------------------------
        'On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        mintMonthIndex = MonthDetails()
        cmbCalMonth.SelectedIndex = mintMonthIndex
        mintCheckPreMonth = cmbCalMonth.SelectedIndex
        Call cmbCalMonth_SelectedIndexChanged(cmbCalMonth, New System.EventArgs())
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub frmMKTMST0002_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '----------------------------
        'Created By     -  Kapil
        'Description    -  If form is deactivated the font of form in the MDIMenu should't be bold
        '----------------------------
        'On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub frmMKTMST0002_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs())
        End If
    End Sub
    Private Sub frmMKTMST0002_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be unloaded
                If Trim(UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) <> "View" Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpMkt.Revert()
                        Me.cmbCalMonth.Enabled = True
                        Me.cbocalYear.Enabled = True
                        cbocalYear.Focus()
                        Call cmbCalMonth_SelectedIndexChanged(cmbCalMonth, New System.EventArgs())
                    Else
                        Me.CmdGrpMkt.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTMST0002_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '----------------------
        'Created By     -  Kapil
        'Description    -  Add form name to the MDIWindow if the form is opened
        '----------------------
        'On Error GoTo ErrHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'Fill Labels From Resource File
        Call FitToClient(Me, Frame1, ctlFormHeader1, CmdGrpMkt) 'To fit the form in the MDI
        Call SelectFinancialYear() 'To select Financial Year From Company Master
        Call AddMonthToCalMonthCmbBox() 'To add Months In the Combo Box
        Call SetDaysDescriptionInSpreadHdr() 'To Set Days Description In Spread Header

        'Added By Ekta Uniyal on 28 Mar 2014 to support multi-unit functionality
        lblHolDay.BackColor = Color.Red
        lblNonHolDay.BackColor = Color.Black
        'End Here

        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub AddYearsToCalYearCmbBox()
        'On Error GoTo ErrHandler
        Dim intCurYear As Short
        Dim arrintYears(20) As Short
        Dim intIndex As Short
        Dim intNo As Short
        intCurYear = Year(GetServerDate) 'Get Current Year
        mstrStartFinYear = CStr(intCurYear) 'initialize
        intIndex = 0
        Dim counter As Short
        counter = intCurYear
        For intNo = counter To (intCurYear + 10)
            cbocalYear.Items.Insert(intIndex, CStr(intCurYear))
            intCurYear = intCurYear + 1
            intIndex = intIndex + 1
        Next
        cbocalYear.SelectedIndex = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub frmMKTMST0002_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Trim(Me.CmdGrpMkt.lbMode) <> "View" Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'Save data before saving
                        Call DispInEditMode("Save")
                        If mblnhasSave Then
                            Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Me.CmdGrpMkt.Revert()
                            gblnCancelUnload = False
                            gblnFormAddEdit = False
                        End If
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    'Set Global VAriable
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                    Me.CmdGrpMkt.Focus()
                End If
            Else
                Me.Dispose()
                Exit Sub
            End If
        End If
        'Checking The Status
        If gblnCancelUnload = True Then eventArgs.Cancel = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0002_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-----------------------------------------
        'Created By     -  Kapil
        'Description    -  At Form Unload Remove Form from the MDIMenu i.e UnChecked the Form
        '-----------------------------------------
        'On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub AddMonthToCalMonthCmbBox()
        '----------------------
        'Created By     -  Kapil
        'Description    -  Add Months Description to the combobox
        'Revised By     -  Nitin Sood (Months are displayed according to calender Year )
        '----------------------
        'On Error GoTo ErrHandler
        VB6.SetItemString(cmbCalMonth, 0, "April")
        VB6.SetItemString(cmbCalMonth, 1, "May")
        VB6.SetItemString(cmbCalMonth, 2, "June")
        VB6.SetItemString(cmbCalMonth, 3, "July")
        VB6.SetItemString(cmbCalMonth, 4, "August")
        VB6.SetItemString(cmbCalMonth, 5, "September")
        VB6.SetItemString(cmbCalMonth, 6, "October")
        VB6.SetItemString(cmbCalMonth, 7, "November")
        VB6.SetItemString(cmbCalMonth, 8, "December")
        VB6.SetItemString(cmbCalMonth, 9, "January")
        VB6.SetItemString(cmbCalMonth, 10, "February")
        VB6.SetItemString(cmbCalMonth, 11, "March")
        cmbCalMonth.SelectedIndex = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub SetDaysDescriptionInSpreadHdr()
        '------------------------------------------
        'Created By     -  Kapil
        'Description    -  Add Months Description to the combobox
        '-----------------------------------------
        'On Error GoTo ErrHandler
        Dim intRowHeight As Short
        With Me.spCalYear
            For intRowHeight = 0 To .MaxRows
                '''.RowHeight(intRowHeight) = 300
                .set_RowHeight(intRowHeight, 300)
            Next intRowHeight
            .Row = 0
            .Col = 1 : .Text = "Sunday"
            .Row = 0
            .Col = 2 : .Text = "Monday"
            .Row = 0
            .Col = 3 : .Text = "Tuesday"
            .Row = 0
            .Col = 4 : .Text = "Wednesday"
            .Row = 0
            .Col = 5 : .Text = "Thursday"
            .Row = 0
            .Col = 6 : .Text = "Friday"
            .Row = 0
            .Col = 7 : .Text = "Saturday"
            .MaxCols = 7
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub DisplayDaysOfMonth()
        '-------------------------------
        'Created By     -  Kapil
        'Description    -  Display Days Of The Month for the Selected Month of the Year
        '-------------------------------
        'On Error GoTo ErrHandler
        Dim strStartYear As String 'Financial Year Start
        Dim strEndYear As String 'Financial Year End
        strStartYear = mstrStartFinYear
        strEndYear = CStr(Val(mstrStartFinYear) + 1)
        Dim strMonth As String 'Get the Month Selected By User from Combo Box
        strMonth = Me.cmbCalMonth.Text
        Select Case strMonth
            Case "April"
                mintMonthIndex = 4
                mintDaysInMonth = 30 'Total Days in the April Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "May"
                mintMonthIndex = 5
                mintDaysInMonth = 31 'Total Days in the May Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "June"
                mintMonthIndex = 6
                mintDaysInMonth = 30 'Total Days in the June Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "July"
                mintMonthIndex = 7
                mintDaysInMonth = 31 'Total Days in the July Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "August"
                mintMonthIndex = 8
                mintDaysInMonth = 31 'Total Days in the August Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "September"
                mintMonthIndex = 9
                mintDaysInMonth = 30 'Total Days in the September Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "October"
                mintMonthIndex = 10
                mintDaysInMonth = 31 'Total Days in the October Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "November"
                mintMonthIndex = 11
                mintDaysInMonth = 30 'Total Days in the November Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "December"
                mintMonthIndex = 12
                mintDaysInMonth = 31 'Total Days in the December Month
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strStartYear, strEndYear)
            Case "January"
                mintMonthIndex = 1
                mintDaysInMonth = 31 'Total Days in the January Month Of Fin. Year Ending
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strEndYear, strEndYear)
            Case "February"
                mintMonthIndex = 2
                If (Val(strEndYear) Mod 4 = 0) And (Val(strEndYear) Mod 100 <> 0) Or (Val(strEndYear) Mod 400 = 0) Then
                    mintDaysInMonth = 29 'Total Days in the February Month Of Fin. Year Ending if it is a Leap Year
                Else
                    mintDaysInMonth = 28 'Total Days in the February Month Of Fin. Year Ending
                End If
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strEndYear, strEndYear)
            Case "March"
                mintMonthIndex = 3
                mintDaysInMonth = 31 'Total Days in the March Month Of Fin. Year Ending
                'Procedure Call to Display the Calender for this Month
                Call DisplayMonthDaysDetails(mintMonthIndex, strEndYear, strEndYear)
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub SelectFinancialYear()
        '-------------------------------
        'Created By     -  Kapil
        'Description    -  Display Current Financial Year From Company Master(Company_Mst)
        '-------------------------------
        'On Error GoTo ErrHandler
        Dim strCalYear As String
        Dim rsCalYear As ClsResultSetDB
        Dim strFinYear As String
        Dim intNo As Short
        Dim intNoRec As Short
        strCalYear = "select Fin_Start_Date,Fin_End_date from Financial_Year_Tb where unit_code = '" & gstrUNITID & "'"
        rsCalYear = New ClsResultSetDB
        rsCalYear.GetResult(strCalYear, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intNoRec = rsCalYear.GetNoRows
        rsCalYear.MoveFirst()
        For intNo = 1 To intNoRec
            strFinYear = VB6.Format(rsCalYear.GetValue("Fin_Start_Date"), gstrDateFormat)
            mstrStartFinYear = VB6.Format(strFinYear, "yyyy")
            strFinYear = VB6.Format(strFinYear, "yyyy") & " - " & CDbl(VB6.Format(strFinYear, "yyyy")) + 1
            cbocalYear.Items.Add(strFinYear)
            rsCalYear.MoveNext()
        Next
        rsCalYear.ResultSetClose()
        'Now Pass Form Label Variables Values According to Curr Year
        mstrStartFinYear = CStr(Year(GetServerDate))
        If (Month(GetServerDate) = 1) Or (Month(GetServerDate) = 2) Or (Month(GetServerDate) = 3) Then
            strFinYear = CDbl(mstrStartFinYear) - 1 & " - " & CStr(CShort(mstrStartFinYear))
            mstrStartFinYear = CStr(Year(GetServerDate) - 1) 'Finacial Year is 1 Less
        Else
            strFinYear = CStr(CShort(mstrStartFinYear)) & " - " & CDbl(mstrStartFinYear) + 1
        End If
        cbocalYear.Text = strFinYear
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Function DisplayMonthDaysDetails(ByRef DaysIndex As Short, ByRef stYear As String, ByRef enYear As String) As Object
        '-------------------------------
        'Created By     -  Kapil
        'Description    -  To Retreive 1st Day of the Month Selected by User
        '-------------------------------
        'On Error GoTo ErrHandler
        Dim intDayIndex As Short
        mstrYear = stYear 'to get Current Finncial Year for the Selected Month
        intDayIndex = CShort(Weekday(CDate("01/" & MonthName(DaysIndex) & "/" & stYear)))
        Select Case intDayIndex
            Case FirstDayOfWeek.Sunday
                'Procedure Call to Display the Next Days Details in the Spread
                'Arguments - Index Of the Day,Days in the Month
                Call DisplayNextWeekDaysDetails(FirstDayOfWeek.Sunday, mintDaysInMonth)
            Case FirstDayOfWeek.Monday
                Call DisplayNextWeekDaysDetails(FirstDayOfWeek.Monday, mintDaysInMonth)
            Case FirstDayOfWeek.Tuesday
                Call DisplayNextWeekDaysDetails(FirstDayOfWeek.Tuesday, mintDaysInMonth)
            Case FirstDayOfWeek.Wednesday
                Call DisplayNextWeekDaysDetails(FirstDayOfWeek.Wednesday, mintDaysInMonth)
            Case FirstDayOfWeek.Thursday
                Call DisplayNextWeekDaysDetails(FirstDayOfWeek.Thursday, mintDaysInMonth)
            Case FirstDayOfWeek.Friday
                Call DisplayNextWeekDaysDetails(FirstDayOfWeek.Friday, mintDaysInMonth)
            Case FirstDayOfWeek.Saturday
                Call DisplayNextWeekDaysDetails(FirstDayOfWeek.Saturday, mintDaysInMonth)
        End Select
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Sub DisplayNextWeekDaysDetails(ByRef DayName As Short, ByRef TotDaysInMonth As Short)
        '-------------------------------
        'Created By     -   Kapil
        'Description    -   To Display the Days Of The Selected Month In the Spread
        'Arguments      -   Day Index i.e Sunday,Mon etc.,Total Days in the Month
        '-------------------------------
        'On Error GoTo ErrHandler
        Dim varCheckDay As Object
        Dim introw As Short
        Dim intRowCount As Short
        Dim intcolcount As Short
        introw = 1 'Start from 1st Row
        intcolcount = DayName 'Index Of the Day i.e 1st Day of the Week for Selected Month
        varCheckDay = 1
        'Procedure to check If Date and Month is Existing in Calendar_Mst the Set the ForeColor
        'of Holidays to Red
        Call CheckExistingMonthHolidays(DayName)
        '---
        With Me.spCalYear
            For intRowCount = introw To .MaxRows
                For intcolcount = intcolcount To .MaxCols
                    Call .SetText(intcolcount, intRowCount, varCheckDay)
                    varCheckDay = varCheckDay + 1
                    If intcolcount = 7 Then 'if Column Count becomes 7 then Move to Next Row
                        intcolcount = 1 'set Starting Column of Next Row
                        If varCheckDay > TotDaysInMonth Then Exit Sub
                        Exit For
                    End If
                    If varCheckDay > TotDaysInMonth Then Exit Sub
                Next intcolcount
            Next intRowCount
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ClearColumnOfSpreadControl(ByRef pArgument As String)
        '-------------------------------
        'Created By     -   Kapil
        'Description    -
        'Arguments      -
        '-------------------------------
        'On Error GoTo ErrHandler
        Dim intRowClear As Short 'Row Count
        Dim intColClear As Short 'Col Count
        Dim varClear As Object
        Select Case pArgument
            Case "Clear" 'To Clear All Columns in the Spread Control
                varClear = ""
                With Me.spCalYear
                    For intRowClear = 1 To .MaxRows
                        For intColClear = 1 To .MaxCols
                            Call .SetText(intColClear, intRowClear, varClear)
                        Next intColClear
                    Next intRowClear
                End With
            Case "View" 'To Make the Text in the Spread Center Align
                With Me.spCalYear
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 7
                    .BlockMode = True
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .BlockMode = False
                End With
            Case "BackColor" 'To set the Back Color and ForeColor of Columns for Sundays
                With Me.spCalYear
                    For intRowClear = 1 To .MaxRows
                        varClear = Nothing
                        Call .GetText(1, intRowClear, varClear)
                        If Len(varClear) > 0 Then
                            .Row = intRowClear
                            .Col = 1
                            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            .ForeColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_RED)
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                            .GridShowHoriz = True
                            .GridShowVert = True
                            .CellBorderStyle = FPSpreadADO.CellBorderStyleConstants.CellBorderStyleBlank
                            .CellBorderType = 0
                        Else
                            .Row = intRowClear
                            .Col = 1
                            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        End If
                    Next intRowClear
                    .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                End With
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub DispInEditMode(ByRef pActiveView As String)
        '-------------------------------
        'Created By     -   Kapil
        'Description    -   To Display Check Box in the Cells of Spread Control from Mon-Sat
        'so that User can Check the Holidays of Month
        'Arguments      -   In view mode User can see the check box in the spread and
        'In Save Mode Save the record in the Calendar_Mst
        '-------------------------------
        'On Error GoTo ErrHandler
        Dim intRCount As Short
        Dim intCCount As Short
        Dim varGetDay As Object
        Dim intDayCount As Short
        Dim intWorkDay As Short
        Dim strInsertSql As String
        Dim strDeleteSQL As String
        Dim strExistDate As String
        Dim strDate As String
        Dim rsExistDate As ClsResultSetDB
        Select Case pActiveView
            Case "View"
                With Me.spCalYear
                    For intRCount = 1 To .MaxRows
                        For intCCount = 1 To .MaxCols
                            varGetDay = Nothing
                            Call .GetText(intCCount, intRCount, varGetDay)
                            If Len(varGetDay) > 0 Then
                                .Row = intRCount
                                .Col = intCCount
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                                .TypeCheckText = varGetDay
                                If .ForeColor = Color.Red Then
                                    .Text = CStr(System.Windows.Forms.CheckState.Checked)
                                Else
                                    .Text = CStr(System.Windows.Forms.CheckState.Unchecked)
                                End If
                            End If
                        Next intCCount
                    Next intRCount
                End With
            Case "Save" 'To Save The Record in the Calendar_Mst
                intDayCount = 1
                mblnhasSave = False
                strInsertSql = ""
                strDeleteSQL = ""
                With Me.spCalYear
                    For intRCount = 1 To .MaxRows
                        For intCCount = 1 To .MaxCols
                            'Select the Text Of Current Column
                            'If Column is not blank then make insert Query for
                            'that Date or if that Date Already Exists then
                            'Delete previous one and Insert Fresh Record
                            varGetDay = Nothing
                            Call .GetText(intCCount, intRCount, varGetDay)
                            If Len(varGetDay) > 0 Then
                                strDate = Trim(Str(intDayCount)) & "/" & Trim(Str(mintMonthIndex)) & "/" & Trim(mstrYear)
                                strExistDate = "select dt from Calendar_mst where"
                                strExistDate = strExistDate & " datepart(yyyy,dt)='" & Trim(mstrYear) & "'"
                                strExistDate = strExistDate & " and datepart(mm,dt) ='" & Trim(Str(mintMonthIndex)) & "'"
                                strExistDate = strExistDate & " And DatePart(dd, dt) = '" & Trim(Str(intDayCount)) & "'"
                                strExistDate &= " and unit_code = '" & gstrUNITID & "'"
                                rsExistDate = New ClsResultSetDB
                                rsExistDate.GetResult(strExistDate, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsExistDate.GetNoRows > 0 Then
                                    strDeleteSQL = strDeleteSQL & "delete from Calendar_mst where "
                                    strDeleteSQL = strDeleteSQL & " datepart(yy,dt)='" & Trim(mstrYear) & "' and datepart(mm,dt) ='" & Trim(Str(mintMonthIndex)) & "'"
                                    strDeleteSQL = strDeleteSQL & " and datepart(dd,dt)='" & Trim(Str(intDayCount)) & "'"
                                    strDeleteSQL &= " and unit_code = '" & gstrUNITID & "'"
                                    rsExistDate.ResultSetClose()
                                    rsExistDate = Nothing
                                End If
                                .Row = intRCount
                                .Col = intCCount : intWorkDay = CShort(.Text)
                                strInsertSql = strInsertSql & "insert Calendar_mst(dt,work_flg,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,unit_code)"
                                strInsertSql = strInsertSql & "  values('" & strDate & "'," & intWorkDay & ",getdate(),'" & Trim(mP_User) & "',getdate(),'" & Trim(mP_User) & "','" & gstrUNITID & "')" & vbCrLf
                                intDayCount = intDayCount + 1
                            End If
                        Next intCCount
                    Next intRCount
                    With mP_Connection
                        .BeginTrans()
                        .Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If strDeleteSQL <> "" Then
                            .Execute(strDeleteSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'Delete Query for Calendar_Mst
                        End If
                        .Execute(strInsertSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'Insert Query for Calendar_Mst
                        .CommitTrans()
                        mblnhasSave = True 'if true then Display Message for
                        'Successful Saving of record in Save Mode of Command Group Button
                    End With
                End With
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub CheckExistingMonthHolidays(ByRef pStartCol As Short)
        '-------------------------------
        'Created By     -   Kapil
        'Description    -   To Check if it is Existing Month in the Calendar_Mst then retreive from Calendar_mst
        'and set the font of those columns Red
        'Arguments      -   pStartCol indicates starting Column in the first Row
        '-------------------------------
        'On Error GoTo ErrHandler
        Dim strExData As String 'To Retreive the Date
        Dim strExDataSql As String 'To Make Select Query
        Dim intRow1 As Short 'Row Counter
        Dim intCol1 As Short 'Column Counter
        Dim intCol As Short 'To Increase the Date
        Dim rsExData As ClsResultSetDB
        intCol = 1 '1st Day of Month
        With Me.spCalYear
            For intRow1 = 1 To .MaxRows
                For intCol1 = pStartCol To .MaxCols
                    'Make Date
                    strExData = Trim(Str(mintMonthIndex)) & "/" & Trim(Str(intCol)) & "/" & Trim(mstrYear)
                    rsExData = New ClsResultSetDB
                    'Select Query
                    strExDataSql = "select work_flg from Calendar_mst where"
                    strExDataSql = strExDataSql & " datepart(yy,dt)='" & mstrYear & "'"
                    strExDataSql = strExDataSql & " and datepart(mm,dt) ='" & Trim(Str(mintMonthIndex)) & "'"
                    strExDataSql = strExDataSql & " And DatePart(dd, dt) = '" & Trim(Str(intCol)) & "' and work_flg = 1"
                    strExDataSql &= " and unit_code = '" & gstrUNITID & "'"
                    rsExData.GetResult(strExDataSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsExData.GetNoRows > 0 Then
                        'if day is a holiday
                        .Row = intRow1
                        .Col = intCol1
                        .ForeColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_RED)
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        rsExData.ResultSetClose()
                    Else 'if day is not holiday
                        .Row = intRow1
                        .Col = intCol1
                        .ForeColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_BLACK)
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        rsExData.ResultSetClose()
                    End If
                    intCol = intCol + 1
                    If intCol1 = 7 Then 'if current column is Last Column then
                        pStartCol = 1 'next column will be 1st Column of next Row
                        Exit For
                    End If
                Next intCol1
            Next intRow1
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub CheckSundaysInDataBase()
        '-------------------------------
        'Created By     -   Kapil
        'Description    -   Check if Sunday shown il the Spread is Not Holiday then Set the Fore Color
        'Of that Sunday Column Black
        '-------------------------------
        Dim intCheckSunRow As Short
        Dim strSunSql As String
        Dim strSunDate As String
        Dim rsSunDay As ClsResultSetDB
        Dim varSunDay As Object
        With Me.spCalYear
            For intCheckSunRow = 1 To .MaxRows
                varSunDay = Nothing
                Call .GetText(1, intCheckSunRow, varSunDay)
                If Len(varSunDay) > 0 Then
                    'Make Date
                    strSunDate = Trim(Str(mintMonthIndex)) & "/" & Trim(Str(varSunDay)) & "/" & Trim(mstrYear)
                    rsSunDay = New ClsResultSetDB
                    'Select Query
                    strSunSql = "select work_flg from Calendar_mst where"
                    strSunSql = strSunSql & " datepart(yy,dt)='" & mstrYear & "'"
                    strSunSql = strSunSql & " and datepart(mm,dt) ='" & Trim(Str(mintMonthIndex)) & "'"
                    strSunSql = strSunSql & " And DatePart(dd, dt) = '" & Trim(Str(varSunDay)) & "' and work_flg = 0"
                    strSunSql &= " and unit_code = '" & gstrUNITID & "'"
                    rsSunDay.GetResult(strSunSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsSunDay.GetNoRows > 0 Then
                        .Row = intCheckSunRow
                        .Col = 1
                        .ForeColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_BLACK)
                        Call .SetText(1, intCheckSunRow, varSunDay)
                    End If
                    rsSunDay.ResultSetClose()
                    rsSunDay = Nothing
                End If
            Next intCheckSunRow
        End With
    End Sub
    Private Function MonthDetails() As Short
        Dim intmonth As Short
        intmonth = Month(GetServerDate())
        Select Case intmonth
            Case 1
                MonthDetails = 9 'January
            Case 2
                MonthDetails = 10
            Case 3
                MonthDetails = 11
            Case 4
                MonthDetails = 0 'April
            Case 5
                MonthDetails = 1
            Case 6
                MonthDetails = 2
            Case 7
                MonthDetails = 3
            Case 8
                MonthDetails = 4
            Case 9
                MonthDetails = 5
            Case 10
                MonthDetails = 6
            Case 11
                MonthDetails = 7
            Case 12
                MonthDetails = 8
        End Select
    End Function
    'Private Function GetServerDate() As Date
    '    Dim objServerDate As ClsResultSetDB 'Class Object
    '    Dim strSQL As String 'Stores the SQL statement
    '    'Build the SQL statement
    '    strSQL = "SELECT  CONVERT(datetime,getdate(),103)"
    '    'Creating the instance
    '    objServerDate = New ClsResultSetDB
    '    With objServerDate
    '        'Open the recordset
    '        Call .GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        'If we have a record, then getting the financial year else exiting
    '        If .GetNoRows <= 0 Then Exit Function
    '        GetServerDate = CDate(VB6.Format(DateValue(.GetValueByNo(0)), gstrDateFormat))
    '        'Closing the recordset
    '        .ResultSetClose()
    '    End With
    '    'Releasing the object
    '    objServerDate = Nothing
    'End Function
    Private Sub CmdGrpMkt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Private Sub CmdGrpMkt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtngrptwo.ButtonClickEventArgs) Handles CmdGrpMkt.ButtonClick
        'On Error GoTo ErrHandler
        '''Converted Code
        Select Case CmdGrpMkt.Mode
            Case "E"
                Call DispInEditMode("View")
                Me.cmbCalMonth.Enabled = False
                Me.cbocalYear.Enabled = False
                Me.spCalYear.Focus()
            Case "S"
                Call DispInEditMode("Save")
                If mblnhasSave Then
                    Call ConfirmWindow(10335, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Me.CmdGrpMkt.Revert()
                    Me.cmbCalMonth.Enabled = True
                    Me.cbocalYear.Enabled = True
                    cbocalYear.Focus()
                    Call cmbCalMonth_SelectedIndexChanged(cmbCalMonth, New System.EventArgs())
                End If
            Case ""
                Call frmMKTMST0002_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case "X"
                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub

End Class