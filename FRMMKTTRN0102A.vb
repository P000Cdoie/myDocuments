Imports System
Imports System.Data.SqlClient
Imports System.Net
Imports System.Collections
Imports System.Data.XmlReadMode
Imports System.Xml
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - e-Way Bill Detail
'Name of Form       - FRMMKTTRN0102A  , 
'Created by         - Swati Pareek
'Created Date       - 01-03-2021
'description        - 
'*********************************************************************************************************************

Public Class FRMMKTTRN0102A
    Public Customer_code As String
    Public Customer_name As String
    Public Invoice_number As String
    Dim _Invoiceno As String
    Dim _IrnNumber As String
    Dim _IrnDate As Date
    Dim _IrnBarcodeString As String
    Public Property Irnnumber() As String
        Get
            Return _IrnNumber
        End Get
        Set(ByVal value As String)
            _IrnNumber = value
        End Set
    End Property
    Public Property IrnDate() As Date
        Get
            Return _IrnDate
        End Get
        Set(ByVal value As Date)
            _IrnDate = value
        End Set
    End Property
    Public Property IrnBarcodeString() As String
        Get
            Return _IrnBarcodeString
        End Get
        Set(ByVal value As String)
            _IrnBarcodeString = value
        End Set
    End Property
    Public Property InvoiceNo() As String
        Get
            Return _Invoiceno
        End Get
        Set(ByVal value As String)
            _Invoiceno = value
        End Set
    End Property

    Private Sub FRMMKTTRN0102A_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            SetBackGroundColorNew(Me, True)
            LblCustCode.Text = Customer_code
            LblCustName.Text = Customer_name
            LblInvoiceNo.Text = Invoice_number


        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Dispose()
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Sub

    Private Sub Btnok_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btnok.Click
        Try
            Get_IRN_QR_Codedetails()
        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Sub
    Private Sub Get_IRN_QR_Codedetails()
        Dim strJSON As String
        Dim res As String
        Dim node As XmlNodeList
        Dim varstr As String
        Dim xmldoc As New XmlDocument()
        Try

            varstr = SqlConnectionclass.ExecuteScalar("SELECT IRN_QR_URL FROM global_flag")
            strJSON = "{""IRNQRCodeString"":""" & TxtQRCode.Text & """}"

            Using clnt As WebClient = New WebClient()
                clnt.Headers.Add("Content-Type:application/json")
                clnt.Headers.Add("Accept:application/json")
                res = clnt.UploadString(varstr, "POST", strJSON)
                xmldoc.LoadXml(res)

                If res.Contains("DocNo") Then
                    node = xmldoc.GetElementsByTagName("DocNo")
                    LblInvNo.Text = node.Item(0).InnerText

                Else
                    MsgBox("Doc No is not present in scaned QR code.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    TxtQRCode.Text = String.Empty
                    LblIRNNo.Text = String.Empty
                    LblIRNDate.Text = String.Empty
                    TxtQRCode.Focus()
                    Exit Sub
                End If

                If res.Contains("Irn") Then
                    node = xmldoc.GetElementsByTagName("Irn")
                    LblIRNNo.Text = node.Item(0).InnerText
                Else
                    MsgBox("Irn String is not present in scaned QR code.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    TxtQRCode.Text = String.Empty
                    LblIRNDate.Text = String.Empty
                    LblInvNo.Text = String.Empty
                    TxtQRCode.Focus()
                    Exit Sub
                End If
                If res.Contains("IrnDt") Then
                    node = xmldoc.GetElementsByTagName("IrnDt")
                    LblIRNDate.Text = node.Item(0).InnerText
                Else
                    MsgBox("Irn Date is not present in scaned QR code.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    TxtQRCode.Text = String.Empty
                    LblInvNo.Text = String.Empty
                    LblIRNNo.Text = String.Empty
                    TxtQRCode.Focus()
                    Exit Sub
                End If
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnFillIRN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFillIRN.Click
        Try
            If Validate_QRdetails() = False Then Return
            Irnnumber = LblIRNNo.Text.Trim
            IrnDate = Convert.ToDateTime(LblIRNDate.Text).ToString("dd-MM-yyyy")
            IrnBarcodeString = TxtQRCode.Text
            InvoiceNo = LblInvNo.Text.Trim
            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Sub

    Private Function Validate_QRdetails() As Boolean
        Try
            If TxtQRCode.Text.Trim = String.Empty Then
                MsgBox("QR Code should not be empty.kindly scan QR Code!", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            End If
            If LblIRNNo.Text = String.Empty Then
                MsgBox("IRN No should not be empty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            End If
            If LblIRNDate.Text = String.Empty Then
                MsgBox("IRN Date should not be empty!", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            End If
            If Not (LblInvoiceNo.Text = LblInvNo.Text) Then
                MsgBox("Scaned Invoice No does not match to selected Invoice No. ", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Function

    Private Sub BtnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClear.Click
        Try
            TxtQRCode.Text = String.Empty
            LblIRNNo.Text = String.Empty
            LblIRNDate.Text = String.Empty
            LblInvNo.Text = String.Empty
        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Sub

   
    Private Sub TxtQRCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQRCode.KeyUp
        Try
            If e.KeyCode = Keys.Enter Then
                Get_IRN_QR_Codedetails()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Sub
End Class