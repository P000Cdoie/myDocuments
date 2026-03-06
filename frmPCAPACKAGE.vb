Imports System.Data.SqlClient
Imports System.Data
'*****************************************************************************
'Revision Date      :   03 May 2011
'Revision By        :   Rajeev Gupta
'                       1092646-Sometimes to Allocate box not gets 100% in CC Alocation form
'*****************************************************************************
'Revision Date      :   03 May 2011
'Revision By        :   Rajeev Gupta
'                       1093551-% column should not be visible in case of EPAY JV in CC Allocation Form
'                                 and if user is type CC Code by hand discription is not come.
'*****************************************************************************************
'   Revised By:       Parveen Kumar
'   Revised On:       06 Feb 2012
'   Issue ID  :       10187151-images are not using in message box.
'*****************************************************************************************
'Modify by alok rai on 14-feb-2012  for change mgmt of feb
'*****************************************************************************************
'   Revised By:       Rajeev Gupta
'   Revised On:       28 May 2012
'   Issue ID  :       10218645 - CC Code get repeated many times while booking purchases
'*****************************************************************************************
'REVISED BY VINOD ON 18 SEP 2012 FOR MULTI UNIT MIGRATION - SMIEL
'ISSUE ID : 10277476 
'*****************************************************************************************

'**************************************************************************
'   Revised by    -   Rajeev Gupta
'   Revised date  -   06 Dec 2012
'   Issue ID      -   10316198 — COST CENTER BUDGET FUNCTIONALITY GL WISE 
'**************************************************************************


Public Class frmPCAPackage
    '-----------------------------------------------------------------------------------------------------------------------------------

    'Modified by    :   PARVEEN KUMAR
    'Modified ON    :   19/05/2011
    '   Modified to support MultiUnit functionality
    '---------------------------------------------------------------------------------------	
    ''Revised By:       Saurav Kumar
    ''Revised On:       04 Oct 2013
    ''Issue ID  :       10462231 - eMpro ISuite Changes
    '***********************************************************************************************************************************
    ''Revised By:       Rajeev Gupta
    ''Revised On:       10 Nov 2014
    ''Issue ID  :       10692112 - Multiple CC Codes for Single Item
    '***********************************************************************************************************************************


    Dim conn As SqlConnection
    Dim sqlcmd As SqlCommand
    Dim sqldr As SqlDataReader
    Dim sqltran As SqlTransaction
    Dim mstrPartCode As String = String.Empty
    Dim mstrMODE As String = String.Empty
    Dim mstrchallan_no As String = String.Empty
    Dim mstrSLCode As String = String.Empty
    Dim mfltGLAmt As Decimal = 0
    Dim mintLineNo As Int16 = 0
    Dim mblnAddViewMode As Boolean = False
    Dim mVoNo As String = String.Empty
    Dim mVoType As String = String.Empty
    Dim mstrPartyCode As String = String.Empty
    Dim mstr_PartyType As String = String.Empty
    Dim iLoop As Integer = 0
    Dim Val_PartCode As Object = Nothing
    Dim Val_PackageCode As Object = Nothing
    Dim Val_UpdatedQty As Object = Nothing

    Dim mStrSql As String = String.Empty
    Dim mblnSaveData As Boolean = False
    Dim mstrErrMsg As String = String.Empty

    Dim mblnCCACC_GLWISE As Boolean = False
    Dim mblnPV_Against_PO As Boolean = False
    Dim mblnDefaultGL As Boolean = False

    Private Enum EnumPCAPACKAGECODE
        VAR_PARTCODE = 1
        VAR_PACKAGECODE = 2
        VAR_MAXPOSSQTY = 3
        VAR_QTY = 4
    End Enum
    Public Property PartCode() As String
        Get
            PartCode = mstrPartCode
        End Get
        Set(ByVal Value As String)
            mstrPartCode = Value
        End Set
    End Property
    Public Property MODE() As String
        Get
            MODE = mstrMODE
        End Get
        Set(ByVal Value As String)
            mstrMODE = Value
        End Set
    End Property
    Public Property challan_no() As String
        Get
            challan_no = mstrchallan_no
        End Get
        Set(ByVal Value As String)
            mstrchallan_no = Value
        End Set
    End Property
    Public Property SLCode() As String
        Get
            SLCode = mstrSLCode
        End Get
        Set(ByVal Value As String)
            mstrSLCode = Value
        End Set
    End Property
    Public Property GLAmount() As Decimal
        Get
            GLAmount = mfltGLAmt
        End Get
        Set(ByVal Value As Decimal)
            mfltGLAmt = Value
        End Set
    End Property
    Public Property intLineNo() As Int16
        Get
            intLineNo = mintLineNo
        End Get
        Set(ByVal Value As Int16)
            mintLineNo = Value
        End Set
    End Property
    Public Property AddViewMode() As Boolean
        Get
            AddViewMode = mblnAddViewMode
        End Get
        Set(ByVal Value As Boolean)
            mblnAddViewMode = Value
        End Set
    End Property
    Public Property VoNo() As String
        Get
            VoNo = mVoNo
        End Get
        Set(ByVal Value As String)
            mVoNo = Value
        End Set
    End Property
    Public Property VoType() As String
        Get
            VoType = mVoType
        End Get
        Set(ByVal Value As String)
            mVoType = Value
        End Set
    End Property
    Public Property PartyCode() As String
        Get
            PartyCode = mstrPartyCode
        End Get
        Set(ByVal Value As String)
            mstrPartyCode = Value
        End Set
    End Property
    Public Property PartyType() As String
        Get
            PartyType = mstr_PartyType
        End Get
        Set(ByVal Value As String)
            mstr_PartyType = Value
        End Set
    End Property

    Public Property CCACC_GLWISE() As Boolean
        Get
            CCACC_GLWISE = mblnCCACC_GLWISE
        End Get
        Set(ByVal Value As Boolean)
            mblnCCACC_GLWISE = Value
        End Set
    End Property
    Public Property PV_Against_PO() As Boolean
        Get
            PV_Against_PO = mblnPV_Against_PO
        End Get
        Set(ByVal Value As Boolean)
            mblnPV_Against_PO = Value
        End Set
    End Property
    Public Property DefaultGL() As Boolean
        Get
            DefaultGL = mblnDefaultGL
        End Get
        Set(ByVal Value As Boolean)
            mblnDefaultGL = Value
        End Set
    End Property

    Private Sub frmFinCCAllocation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim strpartcode As Object
            SetBackGroundColorNew(Me, True)
            Call InitializeSpread()
            
            Call DispalyData(mstrPartCode)


            'txtToAllocate_Per.Text = txtCurrentTot_Per.Text


        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub Calculate_Current_Total()
        Try

            Dim Per_Curr_Tot As Integer = 0
            Dim Val_Curr_Value As Decimal = 0
            Dim Val_Per As Object = Nothing
            Dim Val_Amount As Object = Nothing
            With Me.spd_Data
                For iLoop = 1 To .MaxRows
                    Val_Per = Nothing
                    .GetText(EnumPCAPACKAGECODE.VAR_MAXPOSSQTY, iLoop, Val_Per)
                    Per_Curr_Tot = Per_Curr_Tot + Convert.ToDecimal(Val_Per)
                    Val_Amount = Nothing
                    .GetText(EnumPCAPACKAGECODE.VAR_QTY, iLoop, Val_Amount)
                    Val_Curr_Value = Val_Curr_Value + Convert.ToDecimal(Val_Amount)
                Next
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub DispalyData(ByVal strpartcode As Object)
        Try
            Dim oData As New DataSet()
            Dim iLoop As Integer = 0
            Dim Purchase_OrderNo As String = String.Empty
            If MODE = "ADD_MODE" Then
                mStrSql = " SELECT PART_CODE,PACKAGECODE, QUANTITY , MAXPOSSQTY  FROM TMP_PCA_ITEMSELECTION_PACKAGECODE  WHERE PART_CODE= '" & strpartcode & "' and unit_code ='" & gstrUNITID.Trim & "' and ip_address = '" & gstrIpaddressWinSck.ToString.Trim & "'order by PACKAGECODE"
            End If
            If MODE = "EDIT_MODE" Then
                mStrSql = " SELECT PART_CODE,PACKAGE_CODE as PACKAGECODE, QUANTITY ,MAXPOSSQTY FROM PCA_SCHEDULE_INVOICE_KNOCKOFF  WHERE PART_CODE= '" & strpartcode & "' and unit_code ='" & gstrUNITID.Trim & "' and invoice_no = '" & challan_no & "' order by package_code "
            End If

            Dim oAdpt As New SqlDataAdapter(mStrSql, SqlConnectionclass.GetConnection)
            oAdpt.Fill(oData, "tempTable")
            If oData.Tables(0).Rows.Count > 0 Then
                For iLoop = 1 To oData.Tables(0).Rows.Count
                    AddBlankRow()
                    Call spd_Data.SetText(EnumPCAPACKAGECODE.VAR_PARTCODE, iLoop, oData.Tables(0).Rows(iLoop - 1).Item("part_code").ToString.Trim)
                    Call spd_Data.SetText(EnumPCAPACKAGECODE.VAR_PACKAGECODE, iLoop, oData.Tables(0).Rows(iLoop - 1).Item("PACKAGECODE").ToString.Trim)
                    Call spd_Data.SetText(EnumPCAPACKAGECODE.VAR_MAXPOSSQTY, iLoop, Val(oData.Tables(0).Rows(iLoop - 1).Item("MAXPOSSQTY")))
                    Call spd_Data.SetText(EnumPCAPACKAGECODE.VAR_QTY, iLoop, Val(oData.Tables(0).Rows(iLoop - 1).Item("QUANTITY")))
                Next
            Else
                AddBlankRow()

            End If
            oData.Dispose()


        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        Finally
            GC.Collect()
        End Try

    End Sub

    
    Private Sub InitializeSpread()
        '*********************************************************************************************
        'Author         :       Prashant Rajpal
        'Arguments      :       
        'Return Value   :       None 
        'Description    :       Initialize the Grid Column
        '*********************************************************************************************
        Try
            With spd_Data
                .MaxRows = 0
                .MaxCols = 4
                .Appearance = FPSpreadADO.AppearanceConstants.AppearanceFlat
                .AppearanceStyle = FPSpreadADO.AppearanceStyleConstants.AppearanceStyleClassic
                .Row = 0 : .Col = EnumPCAPACKAGECODE.VAR_PARTCODE : .Text = "Part Code" : .set_ColWidth(EnumPCAPACKAGECODE.VAR_PARTCODE, 15) : .FontBold = True
                .Row = 0 : .Col = EnumPCAPACKAGECODE.VAR_PACKAGECODE : .Text = "Package Code" : .set_ColWidth(EnumPCAPACKAGECODE.VAR_PACKAGECODE, 22) : .FontBold = True
                If PartyCode = "EPAY" Then
                    .Row = 0 : .Col = EnumPCAPACKAGECODE.VAR_MAXPOSSQTY : .Text = "MAX.Poss Quantity" : .set_ColWidth(EnumPCAPACKAGECODE.VAR_MAXPOSSQTY, 7) : .FontBold = True : .ColHidden = True
                Else
                    .Row = 0 : .Col = EnumPCAPACKAGECODE.VAR_MAXPOSSQTY : .Text = "MAX.Poss Quantity" : .set_ColWidth(EnumPCAPACKAGECODE.VAR_MAXPOSSQTY, 7) : .FontBold = True : .ColHidden = False
                End If
                .Row = 0 : .Col = EnumPCAPACKAGECODE.VAR_QTY : .Text = "QUANTITY" : .set_ColWidth(EnumPCAPACKAGECODE.VAR_QTY, 10) : .FontBold = True

                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleDefault
                .set_RowHeight(0, 15)

                .TextTip = FPSpreadADO.TextTipConstants.TextTipFloatingFocusOnly
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub AddBlankRow()

        Try
            If mblnCCACC_GLWISE = True And mblnPV_Against_PO = True And mblnDefaultGL = True Then
                With spd_Data
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_PARTCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = "" : .Lock = True
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_PACKAGECODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = "" : .Lock = True
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_MAXPOSSQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = "0" : .TypeIntegerMin = 0 : .TypeIntegerMax = 100 : .Lock = True
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_QTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = "0.00" : .TypeFloatMax = 9999999999.9999 : .TypeFloatMin = 0 : .Lock = False
                    .set_RowHeight(.MaxRows, 14)
                End With
            Else
                With spd_Data
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_PARTCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = "" : .Lock = True
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_PACKAGECODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = "" : .Lock = True
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_MAXPOSSQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = "0" : .TypeIntegerMin = 0 : .TypeIntegerMax = 100 : .Lock = True
                    .Row = .MaxRows : .Col = EnumPCAPACKAGECODE.VAR_QTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = "0.00" : .TypeFloatMax = 9999999999.9999 : .TypeFloatMin = 0
                    .set_RowHeight(.MaxRows, 14)
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            If ValidData() = True Then
                If SaveData() = False Then
                    MessageBox.Show(mstrErrMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                End If
            Else
                Exit Sub
            End If
            Me.Dispose()
            GC.Collect()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Function ValidData() As Boolean
        Try
            Dim intmaxpossqty As Integer
            Dim intqty As Integer
            Dim intmaxqty As Integer
            Dim strCCCode As String = ""
            Dim iLoop As Integer
            ValidData = True
            For iLoop = 1 To spd_Data.MaxRows
                spd_Data.Row = iLoop
                spd_Data.Col = EnumPCAPACKAGECODE.VAR_MAXPOSSQTY
                intmaxqty = spd_Data.Text

                spd_Data.Col = EnumPCAPACKAGECODE.VAR_QTY
                intqty = spd_Data.Text

                If intqty <= 0 Or intqty > intmaxqty Then
                    MessageBox.Show("Quantity Can't be Zero OR More than Maximum Possible Quantity  ", ResolveResString(100), MessageBoxButtons.OK)
                    spd_Data.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    spd_Data.Focus()
                    ValidData = False
                    Exit Function
                End If

            Next
            Exit Function

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Function

    Private Sub spd_CCData_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles spd_Data.KeyDownEvent, spd_Data.KeyDownEvent
        Try
            Try
                Dim strSQL As String = ""
                Dim strHelp() As String
                If mblnCCACC_GLWISE = True And mblnPV_Against_PO = True And mblnDefaultGL = True Then Exit Sub
                If Me.spd_Data.ActiveCol = EnumPCAPACKAGECODE.VAR_PARTCODE Then
                    If e.keyCode = 112 Then
                        With ctlHelp
                            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                            .ConnectAsUser = gstrCONNECTIONUSER
                            .ConnectThroughDSN = gstrCONNECTIONDSN
                            .ConnectWithPWD = gstrCONNECTIONPASSWORD
                        End With
                        'Chnaging the mouse pointer
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
                        strSQL = "select distinct ccM_ccCode, ccM_ccDesc from Fin_ccmaster where  unit_code = '" & gstrUNITID.Trim & "' and CCM_TRANSTAG = 1"
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "List of CC Code(s)", 1)
                        'Chnaging the mouse pointer
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                With spd_Data
                                    .SetText(EnumPCAPACKAGECODE.VAR_PARTCODE, .ActiveRow, strHelp(0).Trim)
                                    .SetText(EnumPCAPACKAGECODE.VAR_PACKAGECODE, .ActiveRow, strHelp(1).Trim)
                                    .SetText(EnumPCAPACKAGECODE.VAR_MAXPOSSQTY, .ActiveRow, "0")
                                    .SetText(EnumPCAPACKAGECODE.VAR_QTY, .ActiveRow, "0")
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Call Calculate_Current_Total()

                                    Exit Sub
                                End With
                            Else
                                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub spd_CCData_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles spd_Data.KeyPressEvent, spd_Data.KeyPressEvent
        Try
            If e.keyAscii = 39 Then e.keyAscii = 0
            If mblnCCACC_GLWISE = True And mblnPV_Against_PO = True And mblnDefaultGL = True Then Exit Sub
            If spd_Data.ActiveRow = spd_Data.MaxRows And spd_Data.ActiveCol = EnumPCAPACKAGECODE.VAR_QTY And e.keyAscii = 13 Then
                AddBlankRow()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Try
            Me.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try

    End Sub
    Private Function SaveData() As Boolean
        Try
            Dim strSqlqry As String

            With Me.spd_Data
                strSqlqry = ""
                For iLoop = 1 To .MaxRows
                    Val_PartCode = Nothing
                    .GetText(EnumPCAPACKAGECODE.VAR_PARTCODE, iLoop, Val_PartCode)
                    Val_PackageCode = Nothing
                    .GetText(EnumPCAPACKAGECODE.VAR_PACKAGECODE, iLoop, Val_PackageCode)
                    Val_UpdatedQty = Nothing
                    .GetText(EnumPCAPACKAGECODE.VAR_QTY, iLoop, Val_UpdatedQty)
                    If Val_PartCode.ToString.Trim <> "" Then
                        strSqlqry += "UPDATE TMP_PCA_ITEMSELECTION_PACKAGECODE SET QUANTITY= " & Val_UpdatedQty & " WHERE  PART_CODE='" & Val_PartCode & "' AND PACKAGECODE='" & Val_PackageCode & "'AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"

                    End If
                Next
                If strSqlqry.Length > 0 Then
                    SqlConnectionclass.ExecuteNonQuery(strSqlqry)
                End If

            End With

            'STORED INTO COSTCENTRE TABLE

            
            SaveData = True
        Catch ex As Exception
            SaveData = False
            mstrErrMsg = "Records could not saved."
        End Try
    End Function
    Private Sub spd_CCData_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles spd_Data.LeaveCell, spd_Data.LeaveCell
        Try
            
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ResolveResString(115), MessageBoxButtons.OK)
        End Try
    End Sub
End Class