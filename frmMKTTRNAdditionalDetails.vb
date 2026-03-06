Option Strict Off
Option Explicit On
Friend Class frmMKTTRNAdditionalDetails
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'(C) 2001 MIND, All rights reserved
	'
	'File Name          :   frmMKTTRNAdditionalDetails.frm
	'Function           :   Display Additional PO Details
	'Created By         :   Meenu Gupta
	'Created on         :   10, April 2001
	'Revision History   :
    '---------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    04/05/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------
    ''Revised By:       Saurav Kumar
    ''Revised On:       04 Oct 2013
    ''Issue ID  :       10462231 - eMpro ISuite Changes
    '***********************************************************************************************************************************


	Dim m_strSql As String
    Private Sub frmMKTTRNAdditionalDetails_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        SetBackGroundColorNew(Me, True)
        Dim rstAD As New ClsResultSetDB
        rstAD.GetResult(m_pstrSql)
        If rstAD.GetNoRows > 0 Then
            txtSpecialNotes.Text = rstAD.GetValue("Special_Remarks")
            txtPaymentTerms.Text = rstAD.GetValue("Pay_Remarks")
            txtPricesAre.Text = rstAD.GetValue("Price_Remarks")
            txtPkg.Text = rstAD.GetValue("Packing_Remarks")
            txtFreight.Text = rstAD.GetValue("Frieght_Remarks")
            txtTransitInsurance.Text = rstAD.GetValue("Transport_Remarks")
            txtOctroi.Text = rstAD.GetValue("Octorai_Remarks")
            txtModeOfDespatch.Text = rstAD.GetValue("Mode_Despatch")
            txtDeliverySchedule.Text = rstAD.GetValue("Delivery")
        End If
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Me.Dispose()
    End Sub
End Class