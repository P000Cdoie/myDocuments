Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop

Public Class FRMMKTTRN0092
    Inherits System.Windows.Forms.Form

    '********************************************************************************************************
    'Copyright (c)  -   MIND
    'Name of module -   FRMMKTTRN0081.frm
    'Created By     -   Mayur Kumar
    'Created On     -   06-June-2016
    'description    -   File Transfer for FORD(ASN/GRN)
    '               -   FRMMKTTRN0081
    '********************************************************************************************************

    Dim Obj_FSO As Scripting.FileSystemObject
    Dim Upload_FileType As String
    Dim intloopcounter As Int16 = 0
    Dim GRNFile_PickupLocation As String = String.Empty
    Dim ASNFile_PickupLocation As String = String.Empty
    Dim txtFileName As String = String.Empty
    Dim Obj_EX As Excel.Application
    Dim range As Excel.Range

    Private Enum ENUMFILEDETAILS
        VAL_CHECK = 1
        VAL_FILENAME = 2
    End Enum

    Private Sub FRMMKTTRN0081_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, grpBoxGrid, ctlFormHeader, grpboxbtn)
            Me.MdiParent = mdifrmMain
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btn_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rbtn_ASN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_ASN.CheckedChanged
        Try
            If rbtn_ASN.Checked = True Then
                If chk_SelectAll.Checked = True Then
                    chk_SelectAll.Checked = False
                End If
                fsprDtls.MaxRows = 0
                fsprDtls.MaxCols = 0
                FN_FILESELECTION()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rbtn_GRIN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_GRIN.CheckedChanged
        Try
            If rbtn_GRIN.Checked Then
                If chk_SelectAll.Checked = True Then
                    chk_SelectAll.Checked = False
                End If
                fsprDtls.MaxRows = 0
                fsprDtls.MaxCols = 0
                FN_FILESELECTION()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub InitializeSpread_FileDtls()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.fsprDtls
                .MaxRows = 0
                .MaxCols = ENUMFILEDETAILS.VAL_FILENAME
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = ENUMFILEDETAILS.VAL_CHECK : .Text = "SELECT" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMFILEDETAILS.VAL_FILENAME : .Text = "FILE_NAME" : .set_ColWidth(.Col, 50) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .BlockMode = True
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Row2 = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = 1
                .Col2 = .MaxCols
                .Lock = True
                .BlockMode = False
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Public Sub AddRow()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

            With fsprDtls
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = ENUMFILEDETAILS.VAL_CHECK : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMFILEDETAILS.VAL_FILENAME : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
            End With

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FN_FILESELECTION()
        Try
            Dim folderName As String = Nothing
            Dim upldFiles As Scripting.File
            Dim latestFile As String = ""
            Dim Temp As String = Nothing
            Dim strsql As String = String.Empty


            If rbtn_ASN.Checked Then
                strsql = "SELECT TOP 1 ISNULL(ASNFILE_PICKUPLOCATION,'') ASNFILE_PICKUPLOCATION FROM SALES_PARAMETER" & _
                    " where  UNIT_CODE= '" & gstrUNITID & "'"
                ASNFile_PickupLocation = SqlConnectionclass.ExecuteScalar(strsql)
                If Trim(ASNFile_PickupLocation.ToString()) = "" Then
                    MsgBox("ASN FILE Path Does Not Exist")
                    Exit Sub
                End If
                txtFileName = ASNFile_PickupLocation
            End If

            If rbtn_GRIN.Checked Then
                strsql = "SELECT TOP 1 ISNULL(GRNFILE_PICKUPLOCATION,'') GRNFILE_PICKUPLOCATION FROM SALES_PARAMETER" & _
                    " where  UNIT_CODE= '" & gstrUNITID & "'"
                GRNFile_PickupLocation = SqlConnectionclass.ExecuteScalar(strsql)
                If Trim(GRNFile_PickupLocation.ToString()) = "" Then
                    MsgBox("GRIN FILE Path Does Not Exist")
                    Exit Sub
                End If
                txtFileName = GRNFile_PickupLocation
            End If


            If Len(LTrim(RTrim(txtFileName))) > 0 Then
                folderName = txtFileName

                Temp = Mid(StrReverse(txtFileName), 1, InStr(1, StrReverse(txtFileName), "\") - 1)

                Obj_FSO = New Scripting.FileSystemObject
                If InStr(1, Temp, ".") = 0 Then
                    If Obj_FSO.FolderExists(folderName) = False Then
                        MsgBox("Folder Does Not Exist")
                        fsprDtls.MaxRows = 0
                        Exit Sub
                    End If
                Else
                    If Obj_FSO.GetFolder(txtFileName).Files.Count = 0 Then
                        MsgBox("No Files present in the Release Folder.")
                        fsprDtls.MaxRows = 0
                        Exit Sub
                    End If
                    folderName = VB.Left(folderName, Len(folderName) - Len(Temp) - 1)
                End If

                If Obj_FSO.GetFolder(txtFileName).Files.Count = 0 Then
                    MsgBox("No Files present in the Release Folder.")
                    fsprDtls.MaxRows = 0
                    Exit Sub
                End If


                If Obj_FSO.GetFolder(folderName).Files.Count > 0 Then

                    If Trim(latestFile) <> "" Then
                        txtFileName = Obj_FSO.GetFolder(folderName).Path & "\" & latestFile ''& ".csv"
                    End If

                    Call InitializeSpread_FileDtls()
                    ' Mayur
                    For Each upldFiles In Obj_FSO.GetFolder(folderName).Files
                        If Obj_FSO.GetFolder(folderName).Files.Count >= 1 Then
                            If Path.GetExtension(upldFiles.Path) = ".997" Or Path.GetExtension(upldFiles.Path) = ".861" Then
                                latestFile = upldFiles.Path
                                txtFileName = latestFile
                                latestFile = StrReverse(Mid(StrReverse(latestFile), 1, InStr(1, StrReverse(latestFile), "\") - 1))
                                Call AddRow()
                                Me.fsprDtls.Col = ENUMFILEDETAILS.VAL_CHECK
                                Me.fsprDtls.Text = 0
                                Me.fsprDtls.Col = ENUMFILEDETAILS.VAL_FILENAME
                                Me.fsprDtls.Value = latestFile.ToString()
                            End If
                        End If
                    Next

                    ' Mayur
                Else
                    fsprDtls.MaxRows = 0
                End If

            End If


        Catch e As System.IO.FileNotFoundException
            MsgBox("Invalid File Name.", MsgBoxStyle.OkOnly, ResolveResString(100))
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btn_Transfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Transfer.Click
        Try
            Dim file_name As String = String.Empty
            With Me.fsprDtls
                For intloopcounter = 1 To .MaxRows
                    .Row = intloopcounter
                    .Col = ENUMFILEDETAILS.VAL_CHECK
                    If .Value = 1 Then
                        txtFileName = String.Empty
                        .Row = intloopcounter
                        .Col = ENUMFILEDETAILS.VAL_FILENAME
                        file_name = .Text
                        FileFord_Upload(file_name)
                    End If
                Next
                MsgBox("File(s) Transfered Successfully", MsgBoxStyle.Information, "Empro")
                FN_FILESELECTION()
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub chk_SelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_SelectAll.CheckedChanged
        Try

            If chk_SelectAll.Checked = True Then
                With fsprDtls
                    .Col = ENUMFILEDETAILS.VAL_CHECK
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 1
                        End If
                    Next
                End With
            End If

            If chk_SelectAll.Checked = False Then
                With fsprDtls
                    .Col = ENUMFILEDETAILS.VAL_CHECK
                    For intloopcounter = 1 To .MaxRows
                        .Row = intloopcounter
                        If .Lock = False Then
                            .Value = 0
                        End If
                    Next
                End With
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FileFord_Upload(ByRef file_name As String)
        Try
            Dim Obj_FSO As New Scripting.FileSystemObject
            Dim Cell_Data As String = ""
            Dim Row As Object = Nothing
            Dim i As Short = 0
            Dim Data_Row() As String = Nothing
            Dim trans_number As String = ""
            Dim Cell_Data1 As String = ""
            Dim Rev_No As Object = Nothing
            Dim Col As Short = 0
            Dim folderName As String = Nothing
            Dim strSQL As String = String.Empty
            Dim filearray(0) As Object
            Dim upldFileName(0) As Object
            Dim filedate(0) As Object
            Dim latestFile As String = Nothing
            Dim filename As String = String.Empty
            Dim extension As String = String.Empty
            Dim file_extension As String = String.Empty


            If rbtn_ASN.Checked Then
                strSQL = "SELECT TOP 1 ISNULL(ASNFILE_PICKUPLOCATION,'') ASNFILE_PICKUPLOCATION FROM SALES_PARAMETER" & _
                    " where  UNIT_CODE= '" & gstrUNITID & "'"
                ASNFile_PickupLocation = SqlConnectionclass.ExecuteScalar(strSQL)
                txtFileName = ASNFile_PickupLocation + file_name.ToString()
            End If

            If rbtn_GRIN.Checked Then
                strSQL = "SELECT TOP 1 ISNULL(GRNFILE_PICKUPLOCATION,'') GRNFILE_PICKUPLOCATION FROM SALES_PARAMETER" & _
                    " where  UNIT_CODE= '" & gstrUNITID & "'"
                GRNFile_PickupLocation = SqlConnectionclass.ExecuteScalar(strSQL)
                txtFileName = GRNFile_PickupLocation + file_name.ToString()
            End If


            extension = Path.GetFileNameWithoutExtension(txtFileName)
            file_extension = Path.GetExtension(txtFileName)

            If file_extension.Trim.ToString() = ".861" Then
                My.Computer.FileSystem.CopyFile(txtFileName, "C:\GRN\BackUp\" + extension + ".csv", True)

                filename = "C:\GRN\BackUp\" + extension + ".csv"
            End If

            If file_extension.Trim.ToString() = ".997" Then
                My.Computer.FileSystem.CopyFile(txtFileName, "C:\ASN\BackUp\" + extension + ".csv", True)

                filename = "C:\ASN\BackUp\" + extension + ".csv"
            End If

          
            Obj_EX = New Excel.Application
            Obj_EX.Workbooks.Open(Trim(filename))

            Row = 1

            range = Obj_EX.Cells(Row, 1)
            If Not range.Value Is Nothing Then
                Cell_Data = (range.Value.ToString())
            Else
                Cell_Data = ""
            End If

            If Len(Cell_Data) = 0 Then
                MsgBox("There is No Data to Upload ,Please Check Upload File", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If


            If Len(Cell_Data) < 10 Then
                Col = 1 : i = 0
                Cell_Data = ""
                If Not range.Value Is Nothing Then
                    Cell_Data1 = (range.Value.ToString)
                Else
                    Cell_Data1 = ""
                End If
                While Cell_Data1 <> ""
                    Cell_Data = Cell_Data & Cell_Data1 & ","
                    Col = Col + 1
                    range = Obj_EX.Cells(Row, Col)
                    If Not range.Value Is Nothing Then
                        Cell_Data1 = (range.Value.ToString)
                    Else
                        Cell_Data1 = ""
                    End If
                End While
            End If


            Data_Row = Split(Cell_Data, ",", , CompareMethod.Text)
            For i = 0 To UBound(Data_Row)
                Data_Row(i) = Trim(Replace(Data_Row(i), "'", ""))
            Next i

            If file_extension.Trim.ToString() = ".861" Then

                strSQL = "set dateformat dmy; Insert Into ASN_GRN_ACK_UPLOAD(FILE_TYPE,FileName,ACKNOWLEDNO,ASNNO,ACK_DATE,SHOP_TO,SHOP_FROM,INTERNAL_PARTNO,SHIP_QTY,REC_QTY,Ent_Dt,ENT_USER_ID,Upd_Dt,UPD_USER_ID,Unit_Code) " & _
                          " Values ('G','" & file_name.ToString() & "','" & Trim(Data_Row(2)) & "','" & (Trim(Data_Row(1))) & "','" & FN_Date_Conversion_edifact(Trim(Data_Row(0))) & "'," & _
                          "'" & Trim(Data_Row(3)) & "','" & Trim(Data_Row(4)) & "','" & Trim(Data_Row(5)) & "','" & Trim(Data_Row(6)) & "','" & Trim(Data_Row(7)) & "'" & _
                          " , getDate(), " & _
                          " '" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUNITID & "') "

            End If

            If file_extension.Trim.ToString() = ".997" Then
                While Len(Data_Row(7)) < 4
                    Data_Row(7) = "0" + Data_Row(7)
                End While

                strSQL = "set dateformat dmy; Insert Into ASN_GRN_ACK_UPLOAD(FILE_TYPE,FileName,ACKNOWLEDNO,ASNNO,FUNCGROUP,ASN_CONTROL,DOC_TYPE,RESPONSE_CODE,ACK_DATE,ACK_TIME,Ent_Dt,ENT_USER_ID,Upd_Dt,UPD_USER_ID,Unit_Code) " & _
                               " Values ('A','" & file_name.ToString() & "','" & Trim(Data_Row(0)) & "','" & (Trim(Data_Row(1))) & "','" & Trim(Data_Row(2)) & "'," & _
                               "'" & Trim(Data_Row(3)) & "','" & Trim(Data_Row(4)) & "','" & Trim(Data_Row(5)) & "','" & FN_Date_Conversion_edifact(Trim(Data_Row(6))) & "','" & Trim(Data_Row(7)).Substring(0, 2) + ":" + Trim(Data_Row(7)).Substring(2, 2) & "'" & _
                               " , getDate(), " & _
                               " '" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUnitId & "') "


            End If



            SqlConnectionclass.ExecuteNonQuery(strSQL)

            range = Obj_EX.Cells(Row, 1)
            If Not range.Value Is Nothing Then
                Cell_Data = (range.Value.ToString)
            Else
                Cell_Data = ""
            End If

            System.IO.File.Delete(txtFileName)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Function FN_Date_Conversion_edifact(ByRef Cell_Dt As String) As Object

        Try
            Dim T_Month, T_Date, T_Year As String

            Cell_Dt = Replace(Cell_Dt, "'", "")
            If Len(Cell_Dt) >= 5 Then
                T_Date = Mid(Cell_Dt, Len(Cell_Dt) - 1, 2)
                T_Month = Mid(Cell_Dt, Len(Cell_Dt) - 3, 2)
                T_Year = Mid(Cell_Dt, 1, Len(Cell_Dt) - 4)
                If Len(T_Year) = 1 Then
                    T_Year = "200" & T_Year
                ElseIf Len(T_Year) = 2 Then
                    T_Year = "20" & T_Year
                End If
                If IsDate(T_Date & "/" & T_Month & "/" & T_Year) = True Then
                    FN_Date_Conversion_edifact = T_Date & "/" & T_Month & "/" & T_Year
                Else
                    FN_Date_Conversion_edifact = ""
                End If
            Else
                FN_Date_Conversion_edifact = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

End Class