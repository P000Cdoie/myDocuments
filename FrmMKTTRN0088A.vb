'***************************************************************************************
'COPYRIGHT(C)   : MIND   
'FORM NAME      : FRMMKTTRN0088A
'CREATED BY     : VINOD SINGH 
'CREATED DATE   : 12 FEB 2015
'ISSUE ID		: 10737738 - eMPro Vehicle BOM
'***************************************************************************************

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class FRMMKTTRN0088A

#Region "Variables Declaration"

    Private Enum enmCol
        CHK = 1
        Model_Code
        Model_Desc
        Variant_Code
        Variant_Desc
        Volume
    End Enum

#End Region

#Region "Property"
    Dim _SelectedModelVariant As New List(Of ModelVariantStructure)

    Friend Property SelectedModelVariant() As List(Of ModelVariantStructure)
        Get
            Return _SelectedModelVariant
        End Get
        Set(ByVal value As List(Of ModelVariantStructure))
            _SelectedModelVariant = value
        End Set
    End Property

    Dim _CustomerCode As String
    Public Property CustomerCode() As String
        Get
            Return _CustomerCode
        End Get
        Set(ByVal value As String)
            _CustomerCode = value
        End Set
    End Property

    Dim _CustomerName As String
    Public Property CustomerName() As String
        Get
            Return _CustomerName
        End Get
        Set(ByVal value As String)
            _CustomerName = value
        End Set
    End Property

#End Region

#Region "Methods"

    Private Sub AddBlankRow()
        'ADD A NEW ROW IN GRID
        On Error GoTo ErrHandler
        With Me.sprModel
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.Row, 15)
            .Col = enmCol.CHK : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
            .Col = enmCol.Model_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmCol.Model_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmCol.Variant_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmCol.Variant_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            '.Col = enmCol.Usage_Qty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = 0 : .TypeFloatMax = 9999999999.99 : .Lock = True
            .Col = enmCol.Volume : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = 0 : .TypeFloatMax = 9999999999.99
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(CInt(Err.Description), Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub SetGridHeading()
        'SET GRID HEADERS
        On Error GoTo ErrHandler
        With Me.sprModel
            .MaxRows = 0
            .MaxCols = [Enum].GetNames(GetType(enmCol)).Count
            .Row = 0
            .set_RowHeight(0, 20)
            .Col = 0 : .set_ColWidth(0, 3)
            .Col = enmCol.CHK : .Text = " " : .set_ColWidth(enmCol.CHK, 3)
            .Col = enmCol.Model_Code : .Text = "Model Code" : .set_ColWidth(enmCol.Model_Code, 12)
            .Col = enmCol.Model_Desc : .Text = "Model Description" : .set_ColWidth(enmCol.Model_Desc, 20)
            .Col = enmCol.Variant_Code : .Text = "Variant Code" : .set_ColWidth(enmCol.Variant_Code, 12)
            .Col = enmCol.Variant_Desc : .Text = "Variant Description" : .set_ColWidth(enmCol.Variant_Desc, 20)
            '.Col = enmCol.Usage_Qty : .Text = "Usage Qty." : .set_ColWidth(enmCol.Usage_Qty, 8)
            .Col = enmCol.Volume : .Text = "Car Volume" : .set_ColWidth(enmCol.Volume, 8)
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(CInt(Err.Description), Err.Source, Err.Description, mP_Connection)
    End Sub

    Public Sub PopulateModelVariant()

        Try
            sprModel.MaxRows = 0
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "USP_GET_CUSTOMER_MODEL_VARIANT"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 10).Value = _CustomerCode
                    Using dt As DataTable = SqlConnectionclass.GetDataTable(sqlCmd)
                        If dt.Rows.Count > 0 Then
                            With sprModel
                                For Each row As DataRow In dt.Rows
                                    AddBlankRow()
                                    .Row = .MaxRows
                                    .Col = enmCol.CHK : .Value = "0"
                                    .Col = enmCol.Model_Code : .Text = Convert.ToString(row("MODEL_CODE"))
                                    .Col = enmCol.Model_Desc : .Text = Convert.ToString(row("MODEL_DESCRIPTION"))
                                    .Col = enmCol.Variant_Code : .Text = Convert.ToString(row("VARIANT_CODE"))
                                    .Col = enmCol.Variant_Desc : .Text = Convert.ToString(row("VARIANT_DESC"))
                                    .Col = enmCol.Volume : .Value = 0

                                    For Each t As ModelVariantStructure In _SelectedModelVariant
                                        If t.ModelCode.ToUpper = Convert.ToString(row("MODEL_CODE")).ToUpper _
                                        AndAlso t.Variantcode.ToUpper = Convert.ToString(row("VARIANT_CODE")).ToUpper Then
                                            .Col = enmCol.CHK : .Value = "1"
                                            .Col = enmCol.Volume : .Value = t.Volume
                                        End If
                                    Next
                                Next
                            End With
                        End If
                    End Using
                End With
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillSearchCategory()
        Try
            With cboSearchCategory
                .Items.Clear()
                .DataSource = [Enum].GetNames(GetType(enmCol))

            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region "Form & Controls Events"

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        'EXECUTES ON CLICK OF OK BUTTON, SAVED LABELS INFO EITHER IN COLLECTION OR IN DATABASE DEPENDS UPON PARENT SCREEN
        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            With sprModel

                If .MaxRows > 0 Then
                    'VAILIDATE VOLUME QTY
                    For intRow As Integer = 1 To .MaxRows
                        .Row = intRow
                        .Col = enmCol.CHK
                        If .Value = "1" Then
                            .Col = enmCol.Volume
                            If Val(.Value) = 0 Then
                                MsgBox("Car Volume Qty. must be greater than 0 at row#[" & intRow & "]")
                                Return
                            End If
                        End If
                    Next
                    _SelectedModelVariant.Clear()
                    For intRow As Integer = 1 To .MaxRows
                        .Row = intRow
                        .Col = enmCol.CHK
                        If .Value = "1" Then
                            Dim T As New ModelVariantStructure
                            .Col = enmCol.Model_Code : T.ModelCode = .Text.Trim
                            .Col = enmCol.Variant_Code : T.Variantcode = .Text.Trim
                            .Col = enmCol.Volume : T.Volume = Val(.Value)
                            _SelectedModelVariant.Add(T)
                        End If
                    Next
                End If
            End With
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Hide()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        'CANCEL/CLOSE THE SCREEN
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmMKTTRN0088A_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'EXECUTES ON LOADING OF FORM, INITIALIZE CONROLS AND SET DEFAULT VALUES IN CONTROLS.
        Try
            SetBackGroundColorNew(Me, True)
            SetGridHeading()
            lblCustCode.Text = _CustomerCode
            lblCustName.Text = _CustomerName
            PopulateModelVariant()
            FillSearchCategory()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtSearch_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Dim intCounter As Integer
        Dim strText As String
        Dim Col As enmCol
        Col = DirectCast(cboSearchCategory.SelectedIndex + 1, enmCol)
        For intCounter = 1 To sprModel.MaxRows
            With Me.sprModel
                .Row = intCounter : .Col = Col : strText = Trim(.Text)
                If .FontBold = True Then
                    .FontBold = False
                    .Refresh()
                End If
            End With
        Next
        If Len(txtSearch.Text) = 0 Then Exit Sub
        For intCounter = 1 To sprModel.MaxRows
            With Me.sprModel
                .Row = intCounter : .Col = Col : strText = Trim(.Text)
                If Trim(UCase(Mid(strText, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                    .Row = intCounter : .Col = Col : .FontBold = True
                    .Row = intCounter : .Col = Col : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    Exit For
                End If
            End With
        Next
    End Sub

    Private Sub sprModel_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles sprModel.ButtonClicked
        Try
            With sprModel
                If e.col = enmCol.CHK Then
                    .Row = e.row
                    .Col = enmCol.CHK
                    If .Value = "1" Then
                        .Col2 = .MaxCols
                        .Row2 = e.row
                        .BlockMode = True
                        .BackColor = Color.SeaGreen
                        .ForeColor = Color.White
                        .BlockMode = False
                    Else
                        .Col2 = .MaxCols
                        .Row2 = e.row
                        .BlockMode = True
                        .BackColor = Color.White
                        .ForeColor = Color.Black
                        .BlockMode = False
                    End If

                End If

            End With

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

End Class

Friend Structure ModelVariantStructure
    Dim ModelCode As String
    Dim Variantcode As String
    Dim Volume As Double
End Structure