Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.IO
Friend Class frmASNTRN0001
    Inherits System.Windows.Forms.Form

    Dim mintIndex As Short
    Dim StrASN_Type As String = "MITSUBISHI"
    Dim StrIVFileName As String = "HIV122"
    Dim StrITVFileName As String = "HITV122"

    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustHelp.Click
        Try
            Dim StrSql As String
            Dim strHelp() As String
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With

            StrSql = " SELECT CUSTOMER_CODE,CUST_NAME FROM Customer_Mst (NOLOCK) WHERE UNIT_CODE = '" & gstrUNITID & "' AND asn_type = '" & StrASN_Type & "'"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Help", 2)

            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    Me.txtCustomerCode.Text = strHelp(0)
                    Me.LblCustomerName.Text = strHelp(1)
                    Me.txtCustomerCode.Enabled = False
                Else
                    Me.txtCustomerCode.Text = ""
                    Me.LblCustomerName.Text = ""
                    Me.txtCustomerCode.Enabled = False
                    MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                End If
            End If
        Catch ex As Exception
            Call RaiseException(ex)
        End Try
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Try
            Call ShowHelp("underconstruction.htm")
        Catch ex As Exception
            Call RaiseException(ex)
        End Try
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs)
        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("HLPCSTMS0001.htm")
        Catch ex As Exception
            Call RaiseException(ex)
        End Try
    End Sub
    Private Sub frmASNTRN0001_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
            Me.ctlFormHeader1.HeaderString = Mid(Me.ctlFormHeader1.HeaderString, InStr(1, Me.ctlFormHeader1.HeaderString(), "-") + 1, Len(Me.ctlFormHeader1.HeaderString()))
            Call FitToClient(Me, grpMain, ctlFormHeader1, btnDefault, 500)
            Me.MdiParent = mdifrmMain
            btnDefault.Visible = False
            GrpFooter.Left = grpMain.Left

            RefreshScreen()
        Catch ex As Exception
            Dim StrMsg As String
            StrMsg = "frmMKTTRN0096_Load" + vbCrLf + vbCrLf + ex.Message.ToString()
            MessageBox.Show(StrMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmASNTRN0001_Deactivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            Call RaiseException(ex)
        End Try
    End Sub
    Private Sub frmASNTRN0001_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            mdifrmMain.CheckFormName = mintIndex
            System.Windows.Forms.Application.DoEvents()
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            Call RaiseException(ex)
        End Try
    End Sub
    Private Sub frmASNTRN0001_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            frmModules.NodeFontBold(Me.Tag) = False
            mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Catch ex As Exception
            Call RaiseException(ex)
        End Try
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
    End Sub
    Private Sub ASN_MITSUBISHI_FILE()
        Dim StrSql As String
        Dim DtTextFile As DataTable
        Dim SqlCmd As New SqlCommand
        Dim IntTran_Type As Integer

        If OptIV.Checked = True Then
            IntTran_Type = 0
        Else
            IntTran_Type = 1
        End If

        StrSql = " SELECT * FROM [DBO].[UDF_ASN_MITSUBISHI] ('" + gstrUNITID + "'," + IntTran_Type.ToString() + ",'" +
                   txtCustomerCode.Text.Trim() + "','" + Format(dtFromDate.Value, "MM/dd/yyyy") + "','" +
                   Format(dtToDate.Value, "MM/dd/yyyy") + "')"

        With SqlCmd
            .CommandText = StrSql
            .CommandType = CommandType.Text
        End With
        DtTextFile = SqlConnectionclass.GetDataTable(SqlCmd)

        If (DtTextFile.Rows.Count > 0) Then
            If (DtTextFile.Rows(0).Item("FILE_TEXT").ToString().Trim().Length > 0) Then
                Dim StrFilePath As String = DtTextFile.Rows(0).Item("FILE_PATH")

                If (StrFilePath.Trim() = "") Then
                    MsgBox("ASN File path is not defined in Customer Master.", MsgBoxStyle.Critical, ResolveResString(100))
                    Exit Sub
                Else
                    If Directory.Exists(Directory.GetDirectoryRoot(StrFilePath).ToString()) = False Then
                        Directory.CreateDirectory(Directory.GetDirectoryRoot(StrFilePath).ToString())
                    End If
                End If

                If System.IO.File.Exists(StrFilePath) Then
                    Kill(StrFilePath)
                    FileClose(1)
                End If

                Dim StrFile_Text As String = DtTextFile.Rows(0).Item("FILE_TEXT").ToString()
                StrFile_Text = StrFile_Text.Replace("CREATEDATE", DateTime.Now.ToString("yyyy/MM/dd"))
                StrFile_Text = StrFile_Text.Replace("_TIME", DateTime.Now.ToString("HH:mm"))

                FileOpen(1, StrFilePath, OpenMode.Append)
                PrintLine(1, StrFile_Text)
                FileClose(1)

                MsgBox("ASN is Generated successfully.", MsgBoxStyle.Information, ResolveResString(100))
            Else
                MsgBox("Invoice is not exists in given Date Range for customer ", MsgBoxStyle.Critical, ResolveResString(100))
            End If
        Else
            MsgBox("Contact to Database administrator.", MsgBoxStyle.Critical, ResolveResString(100))
        End If
    End Sub
    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Try
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then

                If (txtCustomerCode.Text.Trim() = "") Then
                    MsgBox("Customer can not be blank.", MsgBoxStyle.Critical, ResolveResString(100))
                    Exit Sub
                End If

                ASN_MITSUBISHI_FILE()
                RefreshScreen()
            End If
        Catch ex As Exception
            Call RaiseException(ex)
        End Try
    End Sub
    Private Sub RefreshScreen()
        Me.dtFromDate.Format = DateTimePickerFormat.Custom
        Me.dtFromDate.CustomFormat = gstrDateFormat
        Me.dtFromDate.Value = GetServerDate()
        Me.dtToDate.Format = DateTimePickerFormat.Custom
        Me.dtToDate.CustomFormat = gstrDateFormat
        Me.dtToDate.Value = Me.dtFromDate.Value
        Me.txtCustomerCode.Text = ""
        Me.LblCustomerName.Text = ""
    End Sub
End Class