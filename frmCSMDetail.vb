Imports System.Data.SqlClient

    ''Revised By:       Saurav Kumar
    ''Revised On:       04 Oct 2013
    ''Issue ID  :       10462231 - eMpro ISuite Changes
    '***********************************************************************************************************************************

Public Class frmCSMDetail
    Public mIntInv_no As Integer
    Public mstrFG_item_code As String

    Public Property Inv_no() As Integer
        Get
            Return mIntInv_no
        End Get
        Set(ByVal value As Integer)
            mIntInv_no = value
        End Set
    End Property
    Public Property item_code() As String
        Get
            Return mstrFG_item_code
        End Get
        Set(ByVal value As String)
            mstrFG_item_code = value
        End Set
    End Property

    Private Enum EnumCSMKnocked
        col_FG_item_code = 1
        col_GRN_no = 2
        col_QA_date = 3
        col_CSM_item = 4
        col_knocked_qty = 5
        col_rate = 6
    End Enum
    Private Enum EnumCSMException
        col_FG_item_code = 1
        col_CSM_item
        col_Bal_qty
    End Enum

    Private Sub frmCSMDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetBackGroundColorNew(Me, True)
        initializeCSM_Knocked_Grid()
        initializeCSM_Exception_Grid()
        fill_Knockedoff_Grid()
        fill_Exception_Grid()
        btnClose.Focus()
    End Sub
    Private Sub fill_Knockedoff_Grid()
        Try
            Dim strconString As String = "Server=" & gstrCONNECTIONSERVER & ";Database=" & gstrDatabaseName & ";user=" & gstrCONNECTIONUSER & ";password=" & gstrCONNECTIONPASSWORD & " "
            Dim objcon As New SqlConnection(strconString)
            objcon.Open()
            Dim strsql As String = "SELECT GRN_NO,QA_DATE,CSI_ITEM_CODE,KNOCKED_OFF_QTY,RATE FROM CSM_KNOCKOFF_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INV_NO = " & mIntInv_no & " AND FG_ITEM_CODE = '" & mstrFG_item_code & "'"
            Dim objCom As New SqlCommand(strsql, objcon)
            Dim objReader As SqlDataReader = objCom.ExecuteReader(CommandBehavior.CloseConnection)
            While objReader.Read()
                With fpsCSMKnockedoff
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .SetText(EnumCSMKnocked.col_FG_item_code, .Row, mstrFG_item_code)
                    .SetText(EnumCSMKnocked.col_GRN_no, .Row, objReader("GRN_NO"))
                    .SetText(EnumCSMKnocked.col_QA_date, .Row, VB6.Format(objReader("QA_DATE"), gstrDateFormat))
                    .SetText(EnumCSMKnocked.col_CSM_item, .Row, objReader("CSI_ITEM_CODE"))
                    .SetText(EnumCSMKnocked.col_knocked_qty, .Row, objReader("KNOCKED_OFF_QTY").ToString)
                    .SetText(EnumCSMKnocked.col_rate, .Row, objReader("RATE").ToString)
                End With
            End While
            fpsCSMKnockedoff.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub
    Private Sub fill_Exception_Grid()
        Try
            Dim strconString As String = "Server=" & gstrCONNECTIONSERVER & ";Database=" & gstrDatabaseName & ";user=" & gstrCONNECTIONUSER & ";password=" & gstrCONNECTIONPASSWORD & " "
            Dim objcon As New SqlConnection(strconString)
            objcon.Open()
            Dim strsql As String = "SELECT CSM_ITEM_CODE,REMAINED_QTY FROM CSM_KNOCKOFF_DTL_EXCEPTION WHERE UNIT_CODE='" + gstrUNITID + "' AND INV_NO = " & mIntInv_no & " AND FG_ITEM_CODE = '" & mstrFG_item_code & "'"
            Dim objCom As New SqlCommand(strsql, objcon)
            Dim objReader As SqlDataReader = objCom.ExecuteReader(CommandBehavior.CloseConnection)
            While objReader.Read()
                With fpsCSMException
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .SetText(EnumCSMException.col_FG_item_code, .Row, mstrFG_item_code)
                    .SetText(EnumCSMException.col_CSM_item, .Row, objReader("CSM_ITEM_CODE"))
                    .SetText(EnumCSMException.col_Bal_qty, .Row, objReader("REMAINED_QTY").ToString)
                End With
            End While
            fpsCSMException.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub
    Private Sub initializeCSM_Knocked_Grid()
        With fpsCSMKnockedoff
            .MaxCols = 6
            .MaxRows = 0

            Call .SetText(EnumCSMKnocked.col_FG_item_code, .MaxRows, "FG Item Code")
            Call .SetText(EnumCSMKnocked.col_GRN_no, .MaxRows, "GRIN/SAN No.")
            Call .SetText(EnumCSMKnocked.col_QA_date, .MaxRows, "Date")
            Call .SetText(EnumCSMKnocked.col_CSM_item, .MaxRows, "CSM Item Code")
            Call .SetText(EnumCSMKnocked.col_knocked_qty, .MaxRows, "knocked Off Qty.")
            Call .SetText(EnumCSMKnocked.col_rate, .MaxRows, "Rate")

            .set_RowHeight(0, 20)
            .set_ColWidth(EnumCSMKnocked.col_FG_item_code, 13)
            .set_ColWidth(EnumCSMKnocked.col_GRN_no, 10)
            .set_ColWidth(EnumCSMKnocked.col_QA_date, 9)
            .set_ColWidth(EnumCSMKnocked.col_CSM_item, 13)
            .set_ColWidth(EnumCSMKnocked.col_knocked_qty, 8)
            .set_ColWidth(EnumCSMKnocked.col_rate, 8)

        End With
    End Sub
    Private Sub initializeCSM_Exception_Grid()
        With fpsCSMException
            .MaxRows = 0
            .MaxCols = 3

            Call .SetText(EnumCSMException.col_FG_item_code, .MaxRows, "FG Item Code")
            Call .SetText(EnumCSMException.col_CSM_item, .MaxRows, "CSM Item Code")
            Call .SetText(EnumCSMException.col_Bal_qty, .MaxRows, "Un-knocked Qty.")

            .set_RowHeight(0, 20)
            .set_ColWidth(EnumCSMException.col_FG_item_code, 13)
            .set_ColWidth(EnumCSMException.col_CSM_item, 13)
            .set_ColWidth(EnumCSMException.col_Bal_qty, 9)
        End With
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class