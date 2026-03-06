Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports CrystalDecisions.CrystalReports.Engine
Friend Class frmMKTTRN0019
    Inherits System.Windows.Forms.Form
    '===================================================================================
    ' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   FRMMKTTRN0019.frm
    ' Function          :   Used to Print & Supplementary Invoice
    ' Created By        :   Nisha Rai
    ' Created On        :   03 Nov, 2003
    ' Revision History  :   10 Nov, 2003
    ' Revision History   :   Changes done by Sourabh on 25 oct 2004  for change invoice printing format.
    ' Revised By         : Ashutosh Verma
    ' Revision History   : on 09-01-2006, Issue Id: 16786, Print Supp invoice on new Format.
    '===================================================================================
    ' Revision History  :
    '                   :   Full Roundoff Total Invoice Amount at the time of making master string
    ' Revisised By      :   Parveen Kumar on 07 Mar 2006
    ' Issue Id           :   17290 : Supplementary invoice posting problem (Supp. Invoice - Db. Cr. Mismatch).
    '----------------------------------------------------------------------------------------------------------------
    ' Revision Date     :   28/04/2006
    ' Revision By       :   Davinder Singh
    ' Issue ID          :   17685
    ' Revision History  :   Changes made in the form to call the crystal report for supplymentary Invoice for AIM
    '----------------------------------------------------------------------------------------------------------------
    ' Revision Date     :   01/05/2006
    ' Revision By       :   Davinder Singh
    ' Issue ID          :   17705
    ' Revision History  :   Changes made not to print the Prefix of the Supp Inv No. in case of AIM.
    '-------------------------------------------------------------------------------------------------------
    ' Revision Date     :   25/05/2006
    ' Revision By       :   Davinder Singh
    ' Issue ID          :   17936
    ' Revision History  :   To Separate the code to print Supplementry invoice from general invoice
    '                       by making a new new DLL for it and Give the refrence of this dll to this form
    '----------------------------------------------------------------------------------------------------
    ' Revised By         : Ashutosh Verma
    ' Revision History   : on 21-07-2006, Issue Id: 18326, Consider rounding parameters for taxes in invoice posting.
    '===================================================================================
    ' Revised By         : Ashutosh Verma
    ' Revision History   : On 15-09-2006, Issue Id: 18622,Validate Invoice Number
    '--------------------------------------------------------------------------------
    ' Author              - Davinder Singh
    ' Revision Date       - 26 Apr 2007
    ' Revision History    - To include the Secondary Ecess(SECESS) in Supplementary Invoice
    ' Issue ID            - 19786
    '--------------------------------------------------------------------------------
    ' Revised By          : Davinder Singh
    ' Revision Date       : 01 May 2007
    ' Issue ID            : 19958
    ' History             : To include the concept of MRP in supp inv.
    '                       changes made in this form for posting PKG also
    '--------------------------------------------------------------------------------
    ' Revised By          : Manoj Kr. Vaish
    ' Revision Date       : 30 Aug 2007
    ' Issue ID            : 20956
    ' History             : Inculde Supplementary Invoice Report Printing for SHWOPLA
    '--------------------------------------------------------------------------------
    ' Revised By          : Manoj Kr. Vaish
    ' Revision Date       : 16 Nov 2007
    ' Issue ID            : 21534
    ' History             : Inculde Supplementary Invoice Report Printing for Mate Chennai Units
    '                     : during North South Code Merging Process
    '--------------------------------------------------------------------------------
    'Revised By           : Manoj Kr. Vaish
    'Issue ID             : 21551
    'Revision Date        : 22-Nov-2007
    'History              : Add New Tax VAT with Sale Tax help
    '***********************************************************************************
    'Revised By           : Manoj Kr. Vaish
    'Issue ID             : eMpro-20090309-28488
    'Revision Date        : 09 Mar 2009
    'History              : To pass DSN on prj_Supplementary Invoice printing
    '                       as the DSN was hardcoded in the Dll.
    '***********************************************************************************
    'Revised By           : Manoj Kr. Vaish
    'Revised On           : 16 Mar 2009
    'Issue ID             : eMpro-20090316-28730
    'Revision History     : To get the zone name from sales_parameter
    '------------------------------------------------------------------------------------------------------
    'Revised By           : Siddharth Ranjan
    'Issue ID             : eMpro-20090910-36205
    'Revision Date        : 10-Sep-2009
    'History              : Add Additional VAT functionality
    '-----------------------------------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    08/06/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------------------------------
    '***********************************************************************************
    '============================================================================================
    'Revised By         :   Shalini Singh
    'Revised On         :   23 Sep 2011
    'Reason             :   Ip address and C drive hard coded change
    'issue id           :   10140039
    '============================================================================================
    'Revised By         :   Prashant Rajpal
    'Revised On         :   25 nov 2011
    'Reason             :   Shell Execution Command changed for citrix 
    'issue id           :   10162900
    '============================================================================================
    'Revised By         :   Prashant Rajpal
    'Revised On         :   25 Jan  2012
    'Reason             :   printing location changed 
    'issue id           :   10186031 
    '============================================================================================
    'Revised By         :   Prashant Rajpal
    'Revised On         :   31 jan 2013
    'Reason             :   Supplemetary invoice printing done for MSSl tapukhera
    'issue id           :   10338090  
    '============================================================================================
    'Revised By         :   Prashant Rajpal
    'Revised On         :   10 Aug-2013- 02sep 2013
    'Reason             :   Supplemetary invoice printing done for MSSl tapukhera
    'issue id           :   10229989
    '============================================================================================
    'Revised By         :   Prashant Rajpal
    'Revised On         :   20-aug-2014 
    'Reason             :   Hilex Invoice printing format 
    'issue id           :   10644019 
    '============================================================================================
    ' REVISION DATE     : 25-nov-2014-01 dec 2014
    ' REVISED BY        : PRASHANT RAJPAL
    ' ISSUE ID          : 10717607  
    ' REVISION HISTORY  : MATE MANESAR AND MATE TAPUKARA A4 RELEASED  
    '**********************************************************************************************************************
    ' REVISION DATE     : 26-FEB 2015-27-FEB 2015
    ' REVISED BY        : PRASHANT RAJPAL
    ' ISSUE ID          : 10726518
    ' REVISION HISTORY  : SHIPPING DETAILS DISPLAYED IN REPORTS (CONFIGURABLE )
    '***********************************************************************************************
    ' REVISION DATE     : 06-APR-2015
    ' REVISED BY        : PRASHANT RAJPAL
    ' ISSUE ID          : 10792712 
    ' REVISION HISTORY  : AED CALCULATION IS NOT CONSIDERED IN SUPPLEMENTARY INVOICE. 

    ' REVISION DATE     : 21 JUL 2017
    ' REVISED BY        : ASHISH SHARMA
    ' ISSUE ID          : 101188073
    ' REVISION HISTORY  : GST CHANGES
    '***********************************************************************************************

    'Modify by alok rai on 31-jan-2012 for change mgmt of jan-2012
    Dim mStrCustMst As String
    Dim mintFormIndex As Short
    Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt As Double
    Dim mFrAmt, mGrTotal, mStAmt, mCustmtrl As Double
    Dim mDoc_No As Short
    Dim strCustCode As String 'used in BomCheck() insertupdateAnnex()
    Dim ValidRecord As Boolean
    Dim mAmortization As Double
    Dim mstrMasterString As String 'To store master string for passing to Dr Cr COM
    Dim mstrDetailString As String 'To store detail string for passing to Dr Cr COM
    Dim mstrPurposeCode As String 'To store the Purpose Code which will be used for the fetching of GL and SL
    Dim mblnAddCustomerMaterial As Boolean 'To decide whether to add customer material in basic or not
    Dim mblnSameSeries As Boolean 'To store the flag whether the selected invoice will have same series as others
    Dim mstrReportFilename As String 'To store the report filename
    Dim mblnpostinfin As Boolean
    Dim mblnExciseRoundOFFFlag As Boolean
    Dim mSaleConfNo As Double
    Dim objSuppInvPrinting As prj_SuppInvPrinting.clsSuppInvPrinting
    Dim frmRpt As eMProCrystalReportViewer
    Dim CR As ReportDocument
    Dim strCitrix_Inv_Pronting_Loc As String
    Dim mintnocopies As Short
    '10726518
    Dim mbln_SHIPPING_ADDRESS As Boolean
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        On Error GoTo ErrHandler
        FraInvoicePreview.Visible = False
        objSuppInvPrinting = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Cmdinvoice_ButtonClick1(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        Dim rsInvoiceType As ClsResultSetDB
        Dim strRetVal As String
        Dim objDrCr As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim strInvoiceDate As String
        Dim strRefInvoiceNo As String
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intNoCopies As Short
        Dim strAccountCode As String = String.Empty
        Dim strVoucher As String = String.Empty
        Dim oCmd As ADODB.Command
        Dim strIPAddress As String
        Dim strsql As String

        strIPAddress = gstrIpaddressWinSck

        On Error GoTo Err_Handler
        rsInvoiceType = New ClsResultSetDB
        If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            Me.Close()
            Exit Sub
        Else
            If ValidSelection() = False Then Exit Sub
        End If
        frmRpt = New eMProCrystalReportViewer
        'ISSUE ID 10644019 
        If UCase(Trim(GetPlantName)) = "HILEX" Then
            frmRpt.glblnInvoiceform = True
        Else
            frmRpt.glblnInvoiceform = False
        End If
        'ISSUE ID END  :10644019 
        CR = New ReportDocument
        CR = frmRpt.GetReportDocument
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                If InvoiceGeneration() = True Then
                    '10338090  
                    'If GetPlantName() = "MATM" Then
                    If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.AppStarting)
                        objSuppInvPrinting = New prj_SuppInvPrinting.clsSuppInvPrinting(gstrUNITID, gstrDateFormat)
                        objSuppInvPrinting.DatabaseName = gstrDatabaseName
                        objSuppInvPrinting.DSNName = gstrDSNName
                        objSuppInvPrinting.mstrDSNforInoivcePrint = gstrDSNName
                        objSuppInvPrinting.ConnectionString = gstrCONNECTIONSTRING
                        objSuppInvPrinting.Connection()
                        objSuppInvPrinting.FileName = strCitrix_Inv_Pronting_Loc & "SuppInvoicePrint.txt"
                        objSuppInvPrinting.BCFileName = strCitrix_Inv_Pronting_Loc & "BarCode.txt"
                        objSuppInvPrinting.CompanyName = gstrCOMPANY
                        objSuppInvPrinting.Address1 = gstr_RGN_ADDRESS1
                        objSuppInvPrinting.Address2 = gstr_RGN_ADDRESS2
                        Call objSuppInvPrinting.Print_Invoice_Supp(False, (Me.txtUnitCode).Text, (Me.Ctlinvoice.Text), "")
                        rtbInvoicePreview.LoadFile(objSuppInvPrinting.FileName, RichTextBoxStreamType.PlainText)
                        rtbInvoicePreview.BackColor = System.Drawing.Color.White
                        cmdPrint.Image = My.Resources.ico231.ToBitmap
                        cmdClose.Image = My.Resources.ico217.ToBitmap
                        FraInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1300)
                        FraInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
                        FraInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                        FraInvoicePreview.Top = ctlFormHeader1.Height
                        rtbInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FraInvoicePreview.Height) - 1000)
                        rtbInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FraInvoicePreview.Width) - 200)
                        rtbInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                        rtbInvoicePreview.Top = VB6.TwipsToPixelsY(900)
                        rtbInvoicePreview.RightMargin = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbInvoicePreview.Width) + 5000)
                        shpInvoice.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(FraInvoicePreview.Width) - VB6.PixelsToTwipsX(shpInvoice.Width)) / 2)
                        cmdPrint.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpInvoice.Left) + 100)
                        cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdPrint.Left) + VB6.PixelsToTwipsX(cmdPrint.Width) + 100)
                        cmdPrint.Enabled = True : cmdClose.Enabled = True
                        FraInvoicePreview.Enabled = True : rtbInvoicePreview.Enabled = True : rtbInvoicePreview.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        ReplaceJunkCharacters()
                        FraInvoicePreview.Visible = True
                        FraInvoicePreview.Enabled = True
                        FraInvoicePreview.BringToFront()
                        rtbInvoicePreview.Focus()
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                    Else
                        '10717607
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                        If optInvYes(0).Checked = False Then
                            frmRpt.Show()
                        End If

                        'frmRpt.Show()
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    End If
                Else
                    Exit Sub
                End If
                If chkLockPrintingFlag.CheckState = 1 And optInvYes(0).Checked = True Then
                    Sleep((5000))
                    If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        rsInvoiceType.GetResult("Select Invoice_date from SupplementaryInv_hdr Where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Trim(Ctlinvoice.Text) & "' and Location_code = '" & Trim(txtUnitCode.Text) & "'")
                        strInvoiceDate = getDateForDB(VB6.Format(rsInvoiceType.GetValue("Invoice_date"), gstrDateFormat))
                        rsInvoiceType.ResultSetClose()
                        rsInvoiceType = New ClsResultSetDB
                        rsInvoiceType.GetResult("Select top 1 1 from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "' and Location_code = '" & Trim(txtUnitCode.Text) & "' and supp_invdetail='" & "O" & "'")
                        If rsInvoiceType.GetNoRows > 0 Then
                        Else
                            rsInvoiceType.ResultSetClose()
                            rsInvoiceType = New ClsResultSetDB
                            rsInvoiceType.GetResult("Select RefDoc_no from SupplementaryInv_Dtl Where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Trim(Ctlinvoice.Text) & "'")
                            If rsInvoiceType.GetNoRows > 0 Then
                                rsInvoiceType.MoveFirst()
                                strRefInvoiceNo = rsInvoiceType.GetValue("RefDoc_no")
                                rsInvoiceType.ResultSetClose()
                                rsInvoiceType = New ClsResultSetDB
                                rsInvoiceType.GetResult("Select Invoice_type,sub_category,Account_Code from SalesChallan_dtl where Unit_code='" & gstrUnitId & "' and doc_no = '" & strRefInvoiceNo & "'")
                                strInvoiceType = rsInvoiceType.GetValue("Invoice_type")
                                strInvoiceSubType = rsInvoiceType.GetValue("Sub_category")
                                strAccountCode = rsInvoiceType.GetValue("Account_Code")
                            Else
                                MsgBox("Invalid Invoice No.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                        End If
                        mP_Connection.BeginTrans()
                        mP_Connection.Execute("update SupplementaryInv_hdr set Doc_no = '" & mInvNo & "', Bill_flag = 1 where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("update SupplementaryInv_dtl set Doc_no = '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("Update dbo.SuppCreditAdvise_Dtl set Doc_no = '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If Not mblnSameSeries Then
                            If IsGSTINSAME(strAccountCode) And strInvoiceType = "TRF" Then
                                mP_Connection.Execute("update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Invoice_type = '" & strInvoiceType & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute("update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Invoice_type = '" & strInvoiceType & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        Else
                            If IsGSTINSAME(strAccountCode) And strInvoiceType = "TRF" Then
                                mP_Connection.Execute("update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute("update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                        'Added for Issue ID 1093232 Ends
                        'prashant changed on 05 may 2011 
                        If DataExist("select doc_no from SalesChallan_dtl where location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & mInvNo & "' and Unit_code='" & gstrUNITID & "'") Then
                            MsgBox("Already Saved  with the same Number , Try Again ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                            mP_Connection.RollbackTrans()
                            Exit Sub
                        End If
                        'Added for Issue ID 1093232 Ends
                        'prashant changed ended on 05 may 2011 
                        If mblnpostinfin = True Then
                            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DRG"
                            strRetVal = objDrCr.SetARDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                            'strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            'prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
                            'strRetVal = objDrCr.SetARDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            'prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                            If Mid(strRetVal, 1, 1) = "Y" Then
                                strVoucher = Mid(strRetVal, 16, 12)
                                mP_Connection.Execute("INSERT INTO FIN_AR_INV_MAPPING(VO_NO,UNIT_CODE,INV_NO,SO_NO,ITEM_CODE,HSN_CODE,QTY," & _
                                "Rate,CGST_PER,CGST_AMT,SGST_PER,SGST_AMT,IGST_PER,IGST_AMT,BASIC_AMT,TAX_AMT,TOTAL_AMT,STATUS, " & _
                                "DRCRNOTE, AR_AP,DR_CR) SELECT '" & strVoucher & "','" & gstrUNITID & "','" & mInvNo & "',h.cust_ref,d.item_code,h.hsn_sac_code,0,d.rate_diff," & _
                                "d.CGST_PERCENT,d.DIFF_CGST_AMT,d.sGST_PERCENT,d.DIFF_sGST_AMT,d.IGST_PERCENT,d.DIFF_IGST_AMT ,D.BASIC_AMOUNTDIFF,D.BASIC_AMOUNTDIFF,H.total_amount, " & _
                                "'A','" & strVoucher & "','AR','DR'  from supplementaryinv_hdr h (nolock) ,supplementaryinv_dtl d (nolock)  where h.unit_code=d.unit_code and h.doc_no=d.doc_no AND d.doc_no='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                mP_Connection.Execute("update supplementaryinv_hdr set VOUCHER_NO = '" & strVoucher & "' where Unit_code='" & gstrUNITID & "' and doc_no = '" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            strRetVal = CheckString(strRetVal)
                        Else
                            strRetVal = "Y"
                        End If
                        If Not strRetVal = "Y" Then
                            MsgBox(strRetVal, MsgBoxStyle.Critical, ResolveResString(100))
                            mP_Connection.RollbackTrans()
                            Exit Sub
                        Else
                            '20 dec 2017
                            '20 dec 2017
                            '    mP_Connection.Execute("update tmp_invoiceprint set lorryno= '" & strVoucher & "' where Unit_code='" & gstrUNITID & "' and IP_ADDRESS= '" & strIPAddress & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            CR.DataDefinition.FormulaFields("voucherno").Text = "'" & strVoucher & "'"
                            CR.DataDefinition.FormulaFields("minvno").Text = "'" & mInvNo & "'"
                            frmRpt.Show()
                            mP_Connection.CommitTrans()
                            MsgBox("Supplementary Invoice has been locked successfully with number: " & mInvNo, MsgBoxStyle.Information, ResolveResString(100))
                            Ctlinvoice.Text = ""
                        End If
                    Else
                        frmRpt.Show()
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                If InvoiceGeneration() = True Then
                    'If GetPlantName() = "MATM" Then
                    '    If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                    '        If objSuppInvPrinting Is Nothing Then
                    '            objSuppInvPrinting = New prj_SuppInvPrinting.clsSuppInvPrinting(gstrUNITID, gstrDateFormat)
                    '        End If
                    '        objSuppInvPrinting.DSNName = gstrDSNName
                    '        objSuppInvPrinting.DatabaseName = gstrDatabaseName
                    '        objSuppInvPrinting.mstrDSNforInoivcePrint = gstrDSNName
                    '        objSuppInvPrinting.ConnectionString = gstrCONNECTIONSTRING
                    '        objSuppInvPrinting.Connection()
                    '        objSuppInvPrinting.FileName = strCitrix_Inv_Pronting_Loc & "SuppInvoicePrint.txt"
                    '        objSuppInvPrinting.BCFileName = strCitrix_Inv_Pronting_Loc & "BarCode.txt"
                    '        objSuppInvPrinting.CompanyName = gstrCOMPANY
                    '        objSuppInvPrinting.Address1 = gstr_RGN_ADDRESS1
                    '        objSuppInvPrinting.Address2 = gstr_RGN_ADDRESS2
                    '        Call objSuppInvPrinting.Print_Invoice_Supp(False, (Me.txtUnitCode).Text, (Me.Ctlinvoice.Text), "")
                    '    Else
                    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    '        If optInvYes(0).Checked = True Then
                    '            intMaxLoop = mintnocopies

                    '        Else
                    '            If intNoCopies > 1 Then
                    '                intMaxLoop = mintnocopies
                    '            Else
                    '                intMaxLoop = mintnocopies
                    '            End If
                    '        End If
                    '        For intLoopCounter = 1 To intMaxLoop

                    '            Select Case intLoopCounter
                    '                Case 1
                    '                    If optInvYes(0).Checked = True Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                    '                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                    '                    Else
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER (REPRINT)'"

                    '                    End If
                    '                Case 2
                    '                    If optInvYes(0).Checked = True Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                    '                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                    '                    Else
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER (REPRINT)'"

                    '                    End If
                    '                Case 3
                    '                    If optInvYes(0).Checked = True Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"

                    '                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"

                    '                    Else
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE (REPRINT)'"

                    '                    End If
                    '                Case 4
                    '                    If optInvYes(0).Checked = True Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"

                    '                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"

                    '                    Else
                    '                        CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY (REPRINT)'"
                    '                    End If

                    '            End Select

                    '            frmRpt.SetReportDocument()
                    '            If optInvYes(0).Checked = False Then
                    '                CR.PrintToPrinter(1, False, 0, 0)
                    '            End If

                    '            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    '        Next
                    '    End If
                    'Else
                    '    Exit Sub
                    'End If
                If optInvYes(0).Checked = True Then
                    Sleep((5000))
                    If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        rsInvoiceType.ResultSetClose()
                        rsInvoiceType = New ClsResultSetDB
                        rsInvoiceType.GetResult("Select Invoice_date from SupplementaryInv_hdr Where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Trim(Ctlinvoice.Text) & "' and Location_code = '" & Trim(txtUnitCode.Text) & "'")
                        strInvoiceDate = getDateForDB(VB6.Format(rsInvoiceType.GetValue("Invoice_date"), gstrDateFormat))
                        rsInvoiceType.ResultSetClose()
                        rsInvoiceType = New ClsResultSetDB
                        rsInvoiceType.GetResult("Select RefDoc_no from SupplementaryInv_Dtl Where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Trim(Ctlinvoice.Text) & "'")
                        If rsInvoiceType.GetNoRows > 0 Then
                            rsInvoiceType.MoveFirst()
                            strRefInvoiceNo = rsInvoiceType.GetValue("RefDoc_no")
                            rsInvoiceType.ResultSetClose()
                            rsInvoiceType = New ClsResultSetDB
                            rsInvoiceType.GetResult("Select Invoice_type,sub_category,Account_Code from SalesChallan_dtl where Unit_code='" & gstrUnitId & "' and doc_no = '" & strRefInvoiceNo & "'")
                            strInvoiceType = rsInvoiceType.GetValue("Invoice_type")
                            strInvoiceSubType = rsInvoiceType.GetValue("Sub_category")
                            strAccountCode = rsInvoiceType.GetValue("Account_Code")
                        Else
                            MsgBox("Invalid Invoice No.", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If
                        mP_Connection.BeginTrans()
                        mP_Connection.Execute("update SupplementaryInv_hdr set Doc_no = '" & mInvNo & "', Bill_flag = 1 where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("update SupplementaryInv_dtl set Doc_no = '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("Update dbo.SuppCreditAdvise_Dtl set Doc_no = '" & mInvNo & "' where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If Not mblnSameSeries Then
                            If IsGSTINSAME(strAccountCode) And strInvoiceType = "TRF" Then
                                mP_Connection.Execute("update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Invoice_type = '" & strInvoiceType & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute("update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Invoice_type = '" & strInvoiceType & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        Else
                            If IsGSTINSAME(strAccountCode) And strInvoiceType = "TRF" Then
                                mP_Connection.Execute("update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute("update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                        If mblnpostinfin = True Then
                            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DRG"
                            strRetVal = objDrCr.SetARDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                            'strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            'prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
                            'strRetVal = objDrCr.SetARDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            'prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                            If Mid(strRetVal, 1, 1) = "Y" Then
                                strVoucher = Mid(strRetVal, 16, 12)
                                mP_Connection.Execute("INSERT INTO FIN_AR_INV_MAPPING(VO_NO,UNIT_CODE,INV_NO,SO_NO,ITEM_CODE,HSN_CODE,QTY," & _
                                "Rate,CGST_PER,CGST_AMT,SGST_PER,SGST_AMT,IGST_PER,IGST_AMT,BASIC_AMT,TAX_AMT,TOTAL_AMT,STATUS, " & _
                                "DRCRNOTE, AR_AP,DR_CR) SELECT '" & strVoucher & "','" & gstrUNITID & "','" & mInvNo & "',h.cust_ref,d.item_code,h.hsn_sac_code,0,d.rate_diff," & _
                                "d.CGST_PERCENT,d.DIFF_CGST_AMT,d.sGST_PERCENT,d.DIFF_sGST_AMT,d.IGST_PERCENT,d.DIFF_IGST_AMT ,D.BASIC_AMOUNTDIFF,D.BASIC_AMOUNTDIFF,H.total_amount, " & _
                                "'A','" & strVoucher & "','AR','DR'  from supplementaryinv_hdr h (nolock) ,supplementaryinv_dtl d (nolock)  where h.unit_code=d.unit_code and h.doc_no=d.doc_no AND d.doc_no='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)


                                mP_Connection.Execute("update supplementaryinv_hdr set VOUCHER_NO = '" & strVoucher & "' where Unit_code='" & gstrUNITID & "' and doc_no = '" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            strRetVal = CheckString(strRetVal)
                        Else
                            strRetVal = "Y"
                        End If
                        If Not strRetVal = "Y" Then
                            MsgBox(strRetVal, MsgBoxStyle.Information, ResolveResString(100))
                            mP_Connection.RollbackTrans()
                            Exit Sub
                        Else
                            CR.DataDefinition.FormulaFields("voucherno").Text = "'" & strVoucher & "'"
                            CR.DataDefinition.FormulaFields("minvno").Text = "'" & mInvNo & "'"
                                'CR.PrintToPrinter(1, False, 0, 0)
                            mP_Connection.CommitTrans()
                            MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, ResolveResString(100))
                            cmdPrint.PerformClick()
                            Ctlinvoice.Text = ""
                        End If
                        End If
                        '20 dec 2017
                        If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                            If objSuppInvPrinting Is Nothing Then
                                objSuppInvPrinting = New prj_SuppInvPrinting.clsSuppInvPrinting(gstrUNITID, gstrDateFormat)
                            End If
                            objSuppInvPrinting.DSNName = gstrDSNName
                            objSuppInvPrinting.DatabaseName = gstrDatabaseName
                            objSuppInvPrinting.mstrDSNforInoivcePrint = gstrDSNName
                            objSuppInvPrinting.ConnectionString = gstrCONNECTIONSTRING
                            objSuppInvPrinting.Connection()
                            objSuppInvPrinting.FileName = strCitrix_Inv_Pronting_Loc & "SuppInvoicePrint.txt"
                            objSuppInvPrinting.BCFileName = strCitrix_Inv_Pronting_Loc & "BarCode.txt"
                            objSuppInvPrinting.CompanyName = gstrCOMPANY
                            objSuppInvPrinting.Address1 = gstr_RGN_ADDRESS1
                            objSuppInvPrinting.Address2 = gstr_RGN_ADDRESS2
                            Call objSuppInvPrinting.Print_Invoice_Supp(False, (Me.txtUnitCode).Text, (Me.Ctlinvoice.Text), "")
                        Else
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                            If optInvYes(0).Checked = True Then
                                intMaxLoop = mintnocopies

                            Else
                                If intNoCopies > 1 Then
                                    intMaxLoop = mintnocopies
                                Else
                                    intMaxLoop = mintnocopies
                                End If
                            End If
                            For intLoopCounter = 1 To intMaxLoop

                                Select Case intLoopCounter
                                    Case 1
                                        If optInvYes(0).Checked = True Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                                        ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                                        Else
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER (REPRINT)'"

                                        End If
                                    Case 2
                                        If optInvYes(0).Checked = True Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                                        ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                                        Else
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER (REPRINT)'"

                                        End If
                                    Case 3
                                        If optInvYes(0).Checked = True Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"

                                        ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"

                                        Else
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE (REPRINT)'"

                                        End If
                                    Case 4
                                        If optInvYes(0).Checked = True Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"

                                        ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"

                                        Else
                                            CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY (REPRINT)'"
                                        End If

                                End Select

                                frmRpt.SetReportDocument()
                                'If optInvYes(0).Checked = False Then
                                CR.PrintToPrinter(1, False, 0, 0)
                                '                                End If

                                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                            Next
                        End If
                    Else
                        '20 dec 2017
                        If optInvYes(0).Checked = False Then
                            intMaxLoop = mintnocopies
                        Else
                            If intNoCopies > 1 Then
                                intMaxLoop = mintnocopies
                            Else
                                intMaxLoop = mintnocopies
                            End If
                        End If
                        For intLoopCounter = 1 To intMaxLoop

                            Select Case intLoopCounter
                                Case 1
                                    If optInvYes(0).Checked = True Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                                    Else
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER (REPRINT)'"

                                    End If
                                Case 2
                                    If optInvYes(0).Checked = True Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"

                                    Else
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER (REPRINT)'"

                                    End If
                                Case 3
                                    If optInvYes(0).Checked = True Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"

                                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"

                                    Else
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE (REPRINT)'"

                                    End If
                                Case 4
                                    If optInvYes(0).Checked = True Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"

                                    ElseIf UCase(Trim(GetPlantName)) = "HILEX" Then
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"

                                    Else
                                        CR.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY (REPRINT)'"
                                    End If

                            End Select

                            frmRpt.SetReportDocument()
                            'If optInvYes(0).Checked = False Then
                            CR.PrintToPrinter(1, False, 0, 0)
                            '                                End If

                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        Next
                        '20 dec 2017
                        Exit Sub
                    End If
                    '20 dec 2017

                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE
                If InvoiceGeneration() = True Then
                    frmRpt.ExportToFile()
                    Exit Sub
                End If
        End Select
        frmRpt = Nothing
        Exit Sub
Err_Handler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If Err.Number = 20545 Then
            Resume Next
        Else
            frmRpt = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim varTemp As Object
        Dim strFileName As String
        Dim intNoCopies As Short
        If Len(objSuppInvPrinting.FileName) > 0 Then
            strFileName = objSuppInvPrinting.FileName
            strFileName = strCitrix_Inv_Pronting_Loc & "suppInvoicePrint.txt"
        End If
        If intNoCopies = 0 Then intNoCopies = 1
TypeFileNotFoundCreateRetry:
        For intCount = 1 To intNoCopies
            varTemp = Shell("cmd.exe /c " & strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat " & strFileName, AppWinStyle.Hide)
            Sleep(5000)
            If txtUnitCode.Text = "SUN" Then
               ' varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "TypeToPrn.bat " & gstrLocalCDrive & "BarCodePageFeed.txt", AppWinStyle.Hide)
                varTemp = Shell("cmd.exe /c " & strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat " & strCitrix_Inv_Pronting_Loc & "BarCodePageFeed.txt", AppWinStyle.Hide)
            Else
                'varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "TypeToPrn.bat " & gstrLocalCDrive & "PageFeed.txt", AppWinStyle.Hide)
                varTemp = Shell("cmd.exe /c " & strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat " & strCitrix_Inv_Pronting_Loc & "PageFeed.txt", AppWinStyle.Hide)
            End If
        Next
        Exit Sub
ErrHandler:
        If Err.Number = 53 Then
            'Open App.Path & "\" & "TypeToPrn.bat" For Append As #1
            FileOpen(1, strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat", OpenMode.Append)
            PrintLine(1, "Type %1> prn")
            FileClose(1)
            GoTo TypeFileNotFoundCreateRetry
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        On Error GoTo ErrHandler
        Call ShowCode_Desc("SELECT Unt_CodeID,unt_unitname FROM Gen_UnitMaster WHERE Unt_CodeID='" & gstrUNITID & "' and Unt_Status=1", txtUnitCode)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Ctlinvoice_Change(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles Ctlinvoice.Change
        On Error GoTo ErrHandler
        Ctlinvoice.Text = Replace(Trim(Ctlinvoice.Text), "'", "")
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CtlInvoice_KeyPress(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles Ctlinvoice.KeyPress
        Dim KeyAscii As Short = e.KeyAscii
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If
        DirectCast(Sender, CtlGeneral).KeyPressKeyascii = KeyAscii
    End Sub
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        On Error GoTo ErrHandler
        txtUnitCode.Text = Replace(Trim(txtUnitCode.Text), "'", "")
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
        On Error GoTo ErrHandler
        With txtUnitCode
            .SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        'If Ctrl/Alt/Shift is also pressed
        If Shift <> 0 Then Exit Sub
        'Show the help form when user pressed F1
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())
        If KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strUnitDesc As String
        Dim mobjGLTrans As New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate)
        If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
        strUnitDesc = mobjGLTrans.GetUnit(gstrUNITID, ConnectionString:=gstrCONNECTIONSTRING)
        If CheckString(strUnitDesc) <> "Y" Then
            MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, ResolveResString(100))
            txtUnitCode.Text = ""
            Cancel = True
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub chkLockPrintingFlag_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockPrintingFlag.Enter
        shpLock.Visible = True
    End Sub
    Private Sub chkLockPrintingFlag_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkLockPrintingFlag.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Cmdinvoice.Focus()
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub chkLockPrintingFlag_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockPrintingFlag.Leave
        shpLock.Visible = False
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        Dim strHelp As Object
        On Error GoTo Err_Handler
        Select Case Index
            Case 2
                With Me.Ctlinvoice
                    If optInvYes(0).Checked = True Then
                        strHelp = ShowList(1, .Maxlength, "", "Doc_No", DateColumnNameInShowList("Invoice_Date") & " as Invoice_Date ", "SupplementaryInv_Hdr", " and Doc_No >99000000 and bill_flag = 0 and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    Else
                        strHelp = ShowList(1, .Maxlength, "", "Doc_No", DateColumnNameInShowList("Invoice_Date") & " as Invoice_Date ", "SupplementaryInv_Hdr", " and Doc_No < 99000000 and bill_flag = 1 and cancel_flag = 0 and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    End If
                    .Focus()
                End With
                If Val(strHelp) = -1 Then ' No record
                    Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                Else
                    Me.Ctlinvoice.Text = strHelp
                    gobjDB = New ClsResultSetDB
                    If optInvYes(0).Checked = True Then
                        gobjDB.GetResult("SELECT Doc_NO,Invoice_Date FROM SupplementaryInv_hdr Where Unit_code='" & gstrUNITID & "' and Doc_No >99000000 and bill_flag =0 and Doc_No = '" & strHelp & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    Else
                        gobjDB.GetResult("SELECT Doc_NO,Invoice_date FROM SupplementaryInv_hdr Where Unit_code='" & gstrUNITID & "' and Doc_No <99000000 and bill_flag =1 and Doc_No = '" & strHelp & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    End If
                    If gobjDB.GetNoRows > 0 Then
                    End If
                    gobjDB.ResultSetClose()
                    gobjDB = Nothing
                End If
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Call ShowHelp("HLPMKTTRN0019.htm")
    End Sub
    Private Sub CtlInvoice_KeyUp(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyUpEventArgs) Handles Ctlinvoice.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        On Error GoTo Err_Handler
        If KeyCode = 112 Then
            Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs())
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0019_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0019_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Form_Initialize_Renamed()
        On Error GoTo Err_Handler
        gobjDB = New ClsResultSetDB
        gobjDB.GetResult("SELECT EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin,Excise_RoundOFF FROM sales_parameter where Unit_code='" & gstrUNITID & "' ")
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_RoundOFF")
        gobjDB.ResultSetClose()
        gobjDB = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0019_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub frmMKTTRN0019_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
        Dim rsSalesParameter As New ClsResultSetDB
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
        Call EnableControls(True, Me) 'To Disable controls
        optInvYes(0).Enabled = True : optInvYes(1).Enabled = True : optInvYes(0).Checked = True
        gblnCancelUnload = False
        cmdHelp(2).Image = My.Resources.ico111.ToBitmap

        rsSalesParameter.GetResult("SELECT CITRIX_INV_PRONTING_LOC,REQD_SHIPPING_ADDRESS_SUPPLEMENTARY FROM SALES_PARAMETER where unit_code='" & gstrUNITID & "' ")
        If rsSalesParameter.GetNoRows > 0 Then
            strCitrix_Inv_Pronting_Loc = rsSalesParameter.GetValue("CITRIX_INV_PRONTING_LOC")
            '10726518
            mbln_SHIPPING_ADDRESS = rsSalesParameter.GetValue("REQD_SHIPPING_ADDRESS_SUPPLEMENTARY")
            '10726518
        End If
        rsSalesParameter.ResultSetClose()
        rsSalesParameter = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0019_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValidSelection() As Boolean
        Dim blnInvalidData As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        On Error GoTo Err_Handler
        ValidRecord = False
        lNo = 1
        'Checking if all details have been entered
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(txtUnitCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Location Code"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtUnitCode
        End If
        If Len(Trim(Ctlinvoice.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & ResolveResString(60373)
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = Ctlinvoice
        End If
        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        lNo = lNo + 1
        If blnInvalidData = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            ctlBlank.Focus()
            Exit Function
        End If
        ValidRecord = True
        ValidSelection = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub RefreshForm()
        On Error GoTo ErrHandler
        Call EnableControls(False, Me, True)
        optInvYes(0).Enabled = True : optInvYes(1).Enabled = True : optInvYes(0).Checked = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function InvoiceGeneration() As Boolean
        Dim rstemp As ClsResultSetDB
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim Commissionerate As String
        Dim strsql As String
        Dim strCompMst, DeliveredAdd As String
        Dim strSuffix As String
        Dim strRefInvoiceNo As String
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim strCustomerCode As String = String.Empty
        Dim TinNo As String
        Dim blnPrintTinNo As Boolean
        strRefInvoiceNo = "0"
        Dim oCmd As ADODB.Command
        Dim strIPAddress As String
        Dim strInvoiceDate As String

        On Error GoTo Err_Handler
        rstemp = New ClsResultSetDB
        strIPAddress = gstrIpaddressWinSck

        If Trim(Ctlinvoice.Text) = "" Then InvoiceGeneration = False : Exit Function
        rstemp.GetResult("Select top 1 1 from SupplementaryInv_hdr where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Trim(Ctlinvoice.Text) & "' and Location_code = '" & Trim(txtUnitCode.Text) & "' and supp_invdetail='" & "O" & "'")
        If rstemp.GetNoRows > 0 Then
            strInvoiceType = "Inv"
            strInvoiceSubType = "F"
        Else
            rstemp.ResultSetClose()
            rstemp = New ClsResultSetDB
            rstemp.GetResult("Select RefDoc_no from SupplementaryInv_dtl where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Trim(Ctlinvoice.Text) & "' and Location_code = '" & Trim(txtUnitCode.Text) & "'")
            If rstemp.GetNoRows > 0 Then
                rstemp.MoveFirst()
                strRefInvoiceNo = rstemp.GetValue("RefDoc_no")
                rstemp.ResultSetClose()
                rstemp = New ClsResultSetDB
                rstemp.GetResult("Select Invoice_type,Sub_category,Account_Code From SalesChallan_dtl where Unit_code='" & gstrUnitId & "' and Doc_no = '" & strRefInvoiceNo & "' and Location_code = '" & Trim(txtUnitCode.Text) & "'")
                strInvoiceType = rstemp.GetValue("Invoice_type")
                strInvoiceSubType = rstemp.GetValue("Sub_category")
                strCustomerCode = rstemp.GetValue("Account_Code")
            Else
                MsgBox(" Invalid Invoice No.", MsgBoxStyle.Information, ResolveResString(100))
                Ctlinvoice.Text = "" : Ctlinvoice.Focus() : InvoiceGeneration = False : Exit Function
            End If
        End If
        rstemp.ResultSetClose()
        If UCase(Trim(GetPlantName)) = "HILEX" Then
        Else
            If (IsGSTINSAME(strCustomerCode) = True And strInvoiceType = "TRF") Then
                oCmd = New ADODB.Command
                With oCmd
                    .ActiveConnection = mP_Connection
                    .CommandTimeout = 0
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "PRC_INVOICEPRINTING_MATE"
                    .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                    .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                    .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, "SUP"))
                    .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, "S"))
                    .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                    .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                    MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    oCmd = Nothing
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Function
                End If
                oCmd = Nothing
                strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
            Else
                'If ((UCase(Trim(GetPlantName)) = "MATM" Or UCase(Trim(GetPlantName)) = "MR1")) And CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUnitId + "'")) = False Then
                If (gblnGSTUnit = True) And CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = False Then
                    oCmd = New ADODB.Command
                    With oCmd
                        .ActiveConnection = mP_Connection
                        .CommandTimeout = 0
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "PRC_INVOICEPRINTING_MATE"
                        .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                        .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                        .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, "SUP"))
                        .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, "S"))
                        .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                        .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With
                    If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                        MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                        oCmd = Nothing
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Function
                    End If
                    oCmd = Nothing
                    strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                End If
            End If
        End If

        rstemp = New ClsResultSetDB
        rstemp.GetResult("SELECT inv_GLD_prpsCode,Single_Series,SuppReport_filename,NoCopies  FROM SaleConf WHERE Unit_code='" & gstrUNITID & "' and Invoice_Type='" & strInvoiceType & "' AND Sub_Type ='" & strInvoiceSubType & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ")
        If rstemp.GetNoRows > 0 Then
            mstrPurposeCode = IIf(IsDBNull(rstemp.GetValue("inv_GLD_prpsCode")), "", Trim(rstemp.GetValue("inv_GLD_prpsCode")))
            mblnSameSeries = rstemp.GetValue("Single_Series")
            mintnocopies = rstemp.GetValue("nocopies")
            mstrReportFilename = IIf(IsDBNull(rstemp.GetValue("SuppReport_filename")), "", Trim(rstemp.GetValue("SuppReport_filename")))
            If mstrPurposeCode = "" Then
                MsgBox("Please select a Purpose Code in Sales Configuration", MsgBoxStyle.Information, ResolveResString(100))
                mstrPurposeCode = ""
                InvoiceGeneration = False
                Exit Function
            End If
        Else
            MsgBox("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", MsgBoxStyle.Information, ResolveResString(100))
            InvoiceGeneration = False
            Exit Function
        End If
        rstemp.ResultSetClose()
        '16 feb 2018
        If optInvYes(0).Checked = False Then
            strInvoiceDate = Find_Value("select invoice_date from supplementaryinv_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
            mstrReportFilename = Find_Value("SELECT SuppReport_filename FROM SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type='" & strInvoiceType & "' AND Sub_Type ='" & strInvoiceSubType & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0 ")
        End If
        '16 feb 2018

        strCompMst = "Select Reg_NO,Ecc_No,Range_1,Phone,Fax,Email,PLA_No,LST_No,CST_No,Division,Commissionerate,Invoice_Rule,Tin_no from Company_Mst where Unit_code='" & gstrUNITID & "'"
        rstemp = New ClsResultSetDB
        rstemp.GetResult(strCompMst)
        If rstemp.GetNoRows = 1 Then
            RegNo = rstemp.GetValue("Reg_NO")
            EccNo = rstemp.GetValue("Ecc_No")
            Range = rstemp.GetValue("Range_1")
            Phone = rstemp.GetValue("Phone")
            Fax = rstemp.GetValue("Fax")
            EMail = rstemp.GetValue("Email")
            PLA = rstemp.GetValue("PLA_No")
            UPST = rstemp.GetValue("LST_No")
            CST = rstemp.GetValue("CST_No")
            Division = rstemp.GetValue("Division")
            Commissionerate = rstemp.GetValue("Commissionerate")
            Invoice_Rule = rstemp.GetValue("Invoice_Rule")
            TinNo = rstemp.GetValue("Tin_no")
        End If
        rstemp.ResultSetClose()
        If optInvYes(0).Checked = True Then
            Call InitializeValues()
            rstemp = New ClsResultSetDB
            rstemp.GetResult("Select top 1 1 from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "' and Location_code = '" & Trim(txtUnitCode.Text) & "' and supp_invdetail = 'O'")
            If rstemp.GetNoRows > 0 Then
                mInvNo = CDbl(GenerateInvoiceNo("Inv", "F", getDateForDB(GetServerDate()), ""))
            Else
                mInvNo = CDbl(GenerateInvoiceNo(strInvoiceType, strInvoiceSubType, getDateForDB(GetServerDate()), strCustomerCode))
            End If
            rstemp.ResultSetClose()
            If mblnpostinfin = True Then
                If Not CreateStringForAccounts() Then
                    InvoiceGeneration = False
                    Exit Function
                End If
            End If
        End If
        'If GetPlantName() <> "MATM" Then
        If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = False Then
            rstemp = New ClsResultSetDB
            rstemp.GetResult("Select ConsigneeDetails from Sales_parameter where Unit_code='" & gstrUNITID & "'")
            If rstemp.GetValue("ConsigneeDetails") = False Then
                rstemp.ResultSetClose()
                rstemp = New ClsResultSetDB
                rstemp.GetResult("Select a.* from Customer_Mst a, SupplementaryInv_hdr b where a.unit_code=b.unit_code and a.Customer_code = b.Account_code and b.Doc_No = " & Ctlinvoice.Text & " and b.Location_Code='" & Trim(txtUnitCode.Text) & "' and a.Unit_code='" & gstrUNITID & "'")
                If rstemp.GetNoRows > 0 Then
                    DeliveredAdd = Trim(rstemp.GetValue("Ship_address1"))
                    If Len(Trim(DeliveredAdd)) Then
                        DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rstemp.GetValue("Ship_address2"))
                    Else
                        DeliveredAdd = Trim(rstemp.GetValue("Ship_address2"))
                    End If
                End If
            Else
                rstemp.ResultSetClose()
                rstemp = New ClsResultSetDB
                rstemp.GetResult("Select ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3 from Saleschallan_dtl where Unit_code='" & gstrUNITID & "' and Doc_No = " & strRefInvoiceNo & " and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                If rstemp.GetNoRows > 0 Then
                    DeliveredAdd = Trim(rstemp.GetValue("ConsigneeAddress1"))
                    If Len(Trim(DeliveredAdd)) Then
                        DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rstemp.GetValue("ConsigneeAddress2"))
                    Else
                        DeliveredAdd = Trim(rstemp.GetValue("ConsigneeAddress2"))
                    End If
                    If Len(Trim(DeliveredAdd)) Then
                        DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rstemp.GetValue("ConsigneeAddress3"))
                    Else
                        DeliveredAdd = Trim(rstemp.GetValue("ConsigneeAddress3"))
                    End If
                End If
            End If
            rstemp.ResultSetClose()
            Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2

            If (IsGSTINSAME(strCustomerCode) = True And strInvoiceType = "TRF") Then
                strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
            Else
                'If ((UCase(Trim(GetPlantName)) = "MATM" Or UCase(Trim(GetPlantName)) = "MR1")) Then
                If gblnGSTUnit = True And UCase(GetPlantName) <> "HILEX" Then
                    strsql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                Else
                    strsql = "{SupplementaryInv_hdr.Unit_Code}='" & gstrUNITID & "' and {SupplementaryInv_hdr.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SupplementaryInv_Hdr.Doc_No} =" & Trim(Ctlinvoice.Text)
                End If
            End If

            If mstrReportFilename = "" Then
                MsgBox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, ResolveResString(100))
                InvoiceGeneration = False
                Exit Function
            End If
            If (IsGSTINSAME(strCustomerCode) = True And strInvoiceType = "TRF") Then
                CR.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_A4REPORTS.rpt")
            Else
                CR.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
            End If
            With frmRpt
                CR.DataDefinition.FormulaFields("Registration").Text = "'" & RegNo & "'"
                CR.DataDefinition.FormulaFields("ECC").Text = "'" & EccNo & "'"
                CR.DataDefinition.FormulaFields("Range").Text = "'" & Range & "'"
                CR.DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
                CR.DataDefinition.FormulaFields("CompanyAddress").Text = "'" & Address & "'"
                CR.DataDefinition.FormulaFields("Phone").Text = "'" & Phone & "'"
                CR.DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
                CR.DataDefinition.FormulaFields("EMail").Text = "'" & EMail & "'"
                CR.DataDefinition.FormulaFields("PLA").Text = "'" & PLA & "'"
                CR.DataDefinition.FormulaFields("UPST").Text = "'" & UPST & "'"
                CR.DataDefinition.FormulaFields("CST").Text = "'" & CST & "'"
                CR.DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
                CR.DataDefinition.FormulaFields("commissionerate").Text = "'" & Commissionerate & "'"
                CR.DataDefinition.FormulaFields("InvoiceRule").Text = "'" & Invoice_Rule & "'"
                CR.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
                '10726518
                'CR.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
                If mbln_SHIPPING_ADDRESS = True Then
                    CR.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
                Else
                    CR.DataDefinition.FormulaFields("Address2").Text = "''"
                End If
                '10726518
                If optInvYes(0).Checked = True Then
                    CR.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & CStr(mInvNo) & "'"
                Else
                    CR.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & Ctlinvoice.Text.Trim & "'"
                End If
                blnPrintTinNo = CBool(Find_Value("Select isnull(PrintTinNO,0) as PrintTinNO from sales_parameter where Unit_Code='" & gstrUNITID & "'"))
                If blnPrintTinNo = True Then
                    CR.DataDefinition.FormulaFields("TinNo").Text = "'" & TinNo & "'"
                End If
                If optInvYes(0).Checked = False Then
                    mInvNo =Ctlinvoice.Text  
                    CR.DataDefinition.FormulaFields("minvno").Text = "'" & mInvNo & "'"
                End If
                .ShowPrintButton = True
                .ShowExportButton = True
                .ShowTextSearchButton = True
                .WindowState = FormWindowState.Maximized
                CR.RecordSelectionFormula = strsql
            End With
        End If

        InvoiceGeneration = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function Find_Value(ByRef strField As String) As String
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then
            If IsDBNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub InitializeValues()
        On Error GoTo ErrHandler
        mExDuty = 0 : mInvNo = 0 : mBasicAmt = 0 : msubTotal = 0 : mOtherAmt = 0 : mGrTotal = 0 : mStAmt = 0 : mFrAmt = 0
        mDoc_No = 0 : mCustmtrl = 0 : mAmortization = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub optInvYes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optInvYes.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optInvYes.GetIndex(eventSender)
            Ctlinvoice.Text = ""
        End If
    End Sub
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where Unit_code='" & gstrUNITID & "' and tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            GetTaxGlSl = "N"
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetTaxGlSl = Trim(IIf(IsDBNull(objRecordSet.Fields("tx_glCode").Value), "", objRecordSet.Fields("tx_glCode").Value)) & "»" & Trim(IIf(IsDBNull(objRecordSet.Fields("tx_slCode").Value), "", objRecordSet.Fields("tx_slCode").Value))
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetTaxGlSl = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Private Function CreateStringForAccounts() As Boolean
        Dim objRecordSet As New ADODB.Recordset
        Dim objTmpRecordset As New ADODB.Recordset
        Dim rsSalesInvType As ClsResultSetDB
        Dim strRetVal As String
        Dim strRefInvoiceNo As String
        Dim strInvoiceNo As String
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim strInvoiceDate As String
        Dim strCurrencyCode As String
        Dim dblInvoiceAmt As Double
        Dim dblExchangeRate As Double
        Dim dblBasicAmount As Double
        Dim dblBaseCurrencyAmount As Double
        Dim dblTaxAmt As Double
        Dim strTaxType As String
        Dim strCreditTermsID As String
        Dim strBasicDueDate As String
        Dim strPaymentDueDate As String
        Dim strExpectedDueDate As String
        Dim strCustomerGL As String
        Dim strCustomerSL As String
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim strItemGL As String
        Dim strItemSL As String
        Dim strGlGroupId As String
        Dim dblTaxRate As Double
        Dim varTmp As Object
        Dim dblCCShare As Double
        Dim iCtr As Short
        Dim strCustRef As String
        Dim strParamQuery As String
        Dim rsParameterData As ClsResultSetDB
        Dim blnEOU_FLAG As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnInsIncSTax As Boolean
        Dim blnTotalToolCostRoundOff As Boolean
        Dim blnTCSTax As Boolean
        Dim intBasicRoundOffDecimal As Short
        Dim intSaleTaxRoundOffDecimal As Short
        Dim intExciseRoundOffDecimal As Short
        Dim intSSTRoundOffDecimal As Short
        Dim intTCSRoundOffDecimal As Short
        Dim intToolCostRoundOffDecimal As Short
        Dim blnEcssonCVD As Boolean
        Dim intEcssOnCVDRoundOff As Short
        Dim blnECSSTax As Boolean
        Dim intECSRoundOffDecimal As Short
        Dim blnECSSOnSaleTax As Boolean
        Dim intECSSOnSaleRoundOffDecimal As Short
        Dim blnTurnOverTax As Boolean
        Dim intTurnOverTaxRoundOffDecimal As Short
        Dim blnTotalInvoiceAmount As Boolean
        Dim intTotalInvoiceAmountRoundOffDecimal As Short
        Dim dblInvoiceAmtRoundOff_diff As Double
        Dim blnGSTRoundOff As Boolean
        Dim intGSTRoundOffDecimal As Integer
        Dim strTaxCCCode As String

        strTaxCCCode = ""
        rsSalesInvType = New ClsResultSetDB
        On Error GoTo ErrHandler
        'objRecordSet.Open("SELECT Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,pervalue,Item_Code,Cust_Item_Code,Currency_Code,Rate,Packing,Packing_Amount,Basic_Amount,Accessible_amount,Excise_type,CVD_type,SAD_type,Excise_per,CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,Surcharge_salesTaxType,Surcharge_SalesTax_Per,Surcharge_Sales_Tax_Amount,total_amount,dataPosted,Transport_Type,Vehicle_No,Carriage_Name,SRVDINO,SRVLocation,RejectionPosting,SuppInv_Remarks,remarks,ECESS_Type,ECESS_Per,ECESS_Amount,sales_Quantity,supp_invdetail,TotalInvoiceAmtRoundOff_diff,SECESS_Type,SECESS_Per,SECESS_Amount,MRP,ADDVAT_TYPE,ADDVAT_PER,ADDVAT_AMOUNT  ,AED_TYPE,AED_PER,AED_AMOUNT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,UTGSTTXRT_TYPE,UTGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,CGST_AMT,SGST_AMT,UTGST_AMT,IGST_AMT,CCESS_AMT FROM SupplementaryInv_hdr WHERE Unit_code='" & gstrUnitId & "' and Doc_No='" & Trim(Ctlinvoice.Text) & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        objRecordSet.Open("SELECT Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,pervalue,Item_Code,Cust_Item_Code,Currency_Code,Rate,Packing,Packing_Amount,Basic_Amount,Accessible_amount,Excise_type,CVD_type,SAD_type,Excise_per,CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,Surcharge_salesTaxType,Surcharge_SalesTax_Per,Surcharge_Sales_Tax_Amount,total_amount,dataPosted,Transport_Type,Vehicle_No,Carriage_Name,SRVDINO,SRVLocation,RejectionPosting,SuppInv_Remarks,remarks,ECESS_Type,ECESS_Per,ECESS_Amount,sales_Quantity,supp_invdetail,TotalInvoiceAmtRoundOff_diff,SECESS_Type,SECESS_Per,SECESS_Amount,MRP,ADDVAT_TYPE,ADDVAT_PER,ADDVAT_AMOUNT  ,AED_TYPE,AED_PER,AED_AMOUNT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,UTGSTTXRT_TYPE,UTGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,CGST_AMT=(select sum(DIFF_CGST_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),SGST_AMT=(select sum(DIFF_SGST_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),UTGST_AMT,IGST_AMT=(select sum(DIFF_igst_AMT) from supplementaryinv_dtl sd where sd.unit_code=sh.unit_code and sd.doc_no=sh.doc_no ),CCESS_AMT FROM SupplementaryInv_hdr sh WHERE Unit_code='" & gstrUNITID & "' and Doc_No='" & Trim(Ctlinvoice.Text) & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            MsgBox("Supplementary Invoice details not found", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        strInvoiceNo = CStr(mInvNo)
        strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSet.Fields("Currency_Code").Value), "", objRecordSet.Fields("Currency_Code").Value))
        dblInvoiceAmt = IIf(IsDBNull(objRecordSet.Fields("total_amount").Value), 0, objRecordSet.Fields("total_amount").Value)
        dblInvoiceAmtRoundOff_diff = IIf(IsDBNull(objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value), 0, objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value)
        dblExchangeRate = 1
        strCustCode = Trim(objRecordSet.Fields("Account_Code").Value)
        strCustRef = Trim(IIf(IsDBNull(objRecordSet.Fields("cust_ref").Value), "", objRecordSet.Fields("cust_ref").Value))
        'To Get Refrance Invoice No
        rsSalesInvType.GetResult("Select top 1 1 from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Ctlinvoice.Text & "' and Location_code = '" & Trim(txtUnitCode.Text) & "' and supp_invdetail='" & "O" & "'")
        If rsSalesInvType.GetNoRows > 0 Then
        Else
            rsSalesInvType.ResultSetClose()
            rsSalesInvType = New ClsResultSetDB
            rsSalesInvType.GetResult("Select RefDoc_no from SupplementaryInv_dtl where Unit_code='" & gstrUNITID & "' and Doc_no = '" & Trim(Ctlinvoice.Text) & "' and Location_code = '" & Trim(txtUnitCode.Text) & "'")
            If rsSalesInvType.GetNoRows > 0 Then
                rsSalesInvType.MoveFirst()
                strRefInvoiceNo = rsSalesInvType.GetValue("RefDoc_no")
            Else
                Exit Function
            End If
        End If
        'Retreiving the customer gl, sl and credit term id
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Unit_code='" & gstrUNITID & "' and Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objTmpRecordset.EOF Then
            MsgBox("Customer details not found", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                objTmpRecordset.Close()
                objTmpRecordset = Nothing
            End If
            Exit Function
        End If
        strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_ArCode").Value), "", objTmpRecordset.Fields("Cst_ArCode").Value))
        strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_slCode").Value), "", objTmpRecordset.Fields("Cst_slCode").Value))
        strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
        If strCreditTermsID = "" Then
            MsgBox("Credit Terms not found", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                objTmpRecordset.Close()
                objTmpRecordset = Nothing
            End If
            Exit Function
        End If
        Dim objCreditTerms As New prj_CreditTerm.clsCR_Term_Resolver
        strRetVal = objCreditTerms.RetCR_Term_Dates("", "INV", strCreditTermsID, strInvoiceDate, gstrUNITID, "", "", gstrCONNECTIONSTRING)
        If CheckString(strRetVal) = "Y" Then
            strRetVal = Mid(strRetVal, 3)
            varTmp = Split(strRetVal, "»")
            strBasicDueDate = VB6.Format(varTmp(0), "dd-MMM-yyyy")
            strPaymentDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
            strExpectedDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
        Else
            MsgBox(CheckString(strRetVal), MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                objTmpRecordset.Close()
                objTmpRecordset = Nothing
            End If
            Exit Function
        End If
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
        strParamQuery = strParamQuery & "ECESSOnCVD_Roundoff=isnull(ECESSOnCVD_Roundoff,0),ECESSOnCVDRoundOff_Decimal= isnull(ECESSOnCVDRoundOff_Decimal,0),"
        strParamQuery = strParamQuery & " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal, ECESSOnSaleTax_Roundoff, ECESSOnSaleTaxRoundOff_Decimal, "
        strParamQuery = strParamQuery & " TurnOverTax_RoundOff, TurnOverTaxRoundOff_Decimal, TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,ISNULL(GSTTAX_ROUNDOFF_DECIMAL,0) GSTTAX_ROUNDOFF_DECIMAL,ISNULL(GSTTAX_ROUNDOFF,0) GSTTAX_ROUNDOFF ,suppTotalInvoiceAmount_RoundOff ,suppTotalInvoiceAmountRoundOff_Decimal FROM Sales_Parameter where Unit_code='" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnInsIncSTax = rsParameterData.GetValue("InsInc_SalesTax")
            blnTotalToolCostRoundOff = rsParameterData.GetValue("TotalToolCostRoundoff")
            blnTCSTax = rsParameterData.GetValue("TCSTax_Roundoff")
            intBasicRoundOffDecimal = rsParameterData.GetValue("Basic_Roundoff_decimal")
            intSaleTaxRoundOffDecimal = rsParameterData.GetValue("SalesTax_Roundoff_decimal")
            intExciseRoundOffDecimal = rsParameterData.GetValue("Excise_Roundoff_decimal")
            intSSTRoundOffDecimal = rsParameterData.GetValue("SST_Roundoff_decimal")
            intTCSRoundOffDecimal = rsParameterData.GetValue("TCSTax_Roundoff_decimal")
            intToolCostRoundOffDecimal = rsParameterData.GetValue("TotalToolCostRoundoff_decimal")
            If blnEOU_FLAG = True Then
                blnEcssonCVD = rsParameterData.GetValue("ECESSOnCVD_Roundoff")
                intEcssOnCVDRoundOff = rsParameterData.GetValue("ECESSOnCVDRoundOff_Decimal")
            End If
            blnECSSTax = rsParameterData.GetValue("ECESS_Roundoff")
            intECSRoundOffDecimal = rsParameterData.GetValue("ECESSRoundoff_Decimal")
            blnECSSOnSaleTax = rsParameterData.GetValue("ECESSOnSaleTax_Roundoff")
            intECSSOnSaleRoundOffDecimal = rsParameterData.GetValue("ECESSOnSaleTaxRoundOff_Decimal")
            blnTurnOverTax = rsParameterData.GetValue("TurnOverTax_RoundOff")
            intTurnOverTaxRoundOffDecimal = rsParameterData.GetValue("TurnOverTaxRoundOff_Decimal")
            blnTotalInvoiceAmount = rsParameterData.GetValue("suppTotalInvoiceAmount_RoundOff")
            intTotalInvoiceAmountRoundOffDecimal = rsParameterData.GetValue("suppTotalInvoiceAmountRoundOff_Decimal")
            blnGSTRoundOff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
            intGSTRoundOffDecimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Critical, ResolveResString(100))
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Function
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        If blnTotalInvoiceAmount = False Then
            dblInvoiceAmt = System.Math.Round(dblInvoiceAmt * dblExchangeRate, intTotalInvoiceAmountRoundOffDecimal)
        ElseIf blnTotalInvoiceAmount = True Then
            dblInvoiceAmt = System.Math.Round(dblInvoiceAmt * dblExchangeRate, 0)
        End If
        mstrMasterString = ""
        mstrDetailString = ""
        If gblnGSTUnit = False Then
            mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & strInvoiceDate & "»"
            mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            mstrMasterString = mstrMasterString & dblInvoiceAmt & "»" & dblInvoiceAmt & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
        Else
            mstrMasterString = "M»123»" & strInvoiceDate & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal) & "»0»»»supp. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AR»0»0»" & strInvoiceDate & "»0¦"
        End If
        iCtr = 1
        ''CST/LST/SRT/VAT Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "LST" Or strTaxType = "CST" Or strTaxType = "SRT" Or strTaxType = "VAT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Sales_Tax_Amount").Value)
                If blnISSalesTaxRoundOff = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intSaleTaxRoundOffDecimal)
                ElseIf blnISSalesTaxRoundOff = True Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SalesTax_Per").Value), 0, objRecordSet.Fields("SalesTax_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '********* FOR ADITIONAL VAT POSTING *********************
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ADDVAT_TYPE").Value), "", objRecordSet.Fields("ADDVAT_TYPE").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ADDVAT_TYPE").Value), "", objRecordSet.Fields("ADDVAT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "ADVAT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ADDVAT_AMOUNT").Value), 0, objRecordSet.Fields("ADDVAT_AMOUNT").Value)
                If blnISSalesTaxRoundOff = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intSaleTaxRoundOffDecimal)
                ElseIf blnISSalesTaxRoundOff = True Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("ADDVAT_PER").Value), 0, objRecordSet.Fields("ADDVAT_PER").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '**********************END OF ADITIONAL VAT POSTING *********************
        'ECS Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "ECS" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ECESS_Amount").Value), 0, objRecordSet.Fields("ECESS_Amount").Value)
                If blnECSSTax = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intECSRoundOffDecimal)
                Else
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("ECESS_Per").Value), 0, objRecordSet.Fields("ECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        ''SECESS Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "ECSSH" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SECESS_Amount").Value), 0, objRecordSet.Fields("SECESS_Amount").Value)
                If blnECSSTax = False Then
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intECSRoundOffDecimal)
                Else
                    dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
                End If
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SECESS_Per").Value), 0, objRecordSet.Fields("SECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If

        '10792712 
        ''AED 
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("AED_TYPE").Value), "", objRecordSet.Fields("AED_TYPE").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE Unit_code='" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("AED_TYPE").Value), "", objRecordSet.Fields("AED_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "AED" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("AED_AMOUNT").Value), 0, objRecordSet.Fields("AED_AMOUNT").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("AED_PER").Value), 0, objRecordSet.Fields("AED_PER").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '10792712 ENDED 

        ''Packing Posting
        Dim curPacking_per As Decimal
        curPacking_per = IIf(IsDBNull(objRecordSet.Fields("Packing").Value), 0, objRecordSet.Fields("Packing").Value)
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Packing_Amount").Value), 0, objRecordSet.Fields("Packing_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If curPacking_per > 0 Then
            strRetVal = GetTaxGlSl("PKT")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for PACKING", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If dblBaseCurrencyAmount > 0 Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»PKT»0»" & "»»" & curPacking_per & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                iCtr = iCtr + 1
            End If
        End If
        'SST Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value)
        If blnISSurChargeTaxRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intSSTRoundOffDecimal)
        ElseIf blnISSurChargeTaxRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("Surcharge_SalesTax_Per").Value), 0, objRecordSet.Fields("Surcharge_SalesTax_Per").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("SST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for SST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            iCtr = iCtr + 1
        End If
        '101188073
        'CGST
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CGST_AMT").Value), 0, objRecordSet.Fields("CGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CGST_PERCENT").Value), 0, objRecordSet.Fields("CGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("CGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for CGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»CGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                     dblTaxAmt & "»»CGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If

            iCtr = iCtr + 1
        End If
        'SGST
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SGST_AMT").Value), 0, objRecordSet.Fields("SGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SGST_PERCENT").Value), 0, objRecordSet.Fields("SGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("SGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for SGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»sGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If
            iCtr = iCtr + 1
        End If
        'UTGST
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("UTGST_AMT").Value), 0, objRecordSet.Fields("UTGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("UTGST_PERCENT").Value), 0, objRecordSet.Fields("UTGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("UTGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for UTGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»UTGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                        dblTaxAmt & "»»UTGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If



            iCtr = iCtr + 1
        End If
        'IGST
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("IGST_AMT").Value), 0, objRecordSet.Fields("IGST_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("IGST_PERCENT").Value), 0, objRecordSet.Fields("IGST_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("IGST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for IGST", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If gblnGSTUnit = False Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»IGST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»IGST for supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If

            iCtr = iCtr + 1
        End If
        'CCESS
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CCESS_AMT").Value), 0, objRecordSet.Fields("CCESS_AMT").Value)
        If blnGSTRoundOff = False Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intGSTRoundOffDecimal)
        ElseIf blnGSTRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
        End If
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_PERCENT").Value), 0, objRecordSet.Fields("COMPENSATION_CESS_PERCENT").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("GSTEC")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for COMP.CESS", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»GSTEC»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            iCtr = iCtr + 1
        End If
        '101188073
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT sum(SupplementaryInv_dtl.basic_amountdiff )as basic_amount , item_mst.GlGrp_code ,SupplementaryInv_dtl.totalexciseamount FROM SupplementaryInv_Hdr, item_mst, SupplementaryInv_dtl  WHERE SupplementaryInv_Hdr.Unit_code=item_mst.Unit_code and SupplementaryInv_Hdr.unit_code=SupplementaryInv_dtl.unit_code and SupplementaryInv_Hdr.doc_no =SupplementaryInv_dtl.doc_no  and SupplementaryInv_Hdr.Doc_No='" & Trim(Ctlinvoice.Text) & "' and SupplementaryInv_Hdr.Item_Code=item_mst.Item_Code and SupplementaryInv_Hdr.Location_Code='" & Trim(txtUnitCode.Text) & "' and SupplementaryInv_Hdr.Unit_Code='" & gstrUNITID & "' group by item_mst.GlGrp_code,SupplementaryInv_dtl.totalexciseamount ")
        If objRecordSet.EOF Then
            MsgBox("Item details not found.", MsgBoxStyle.Information, ResolveResString(100))
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        While Not objRecordSet.EOF
            strGlGroupId = Trim(IIf(IsDBNull(objRecordSet.Fields("GlGrp_code").Value), "", objRecordSet.Fields("GlGrp_code").Value))
            'Basic Amount Posting
            dblBasicAmount = IIf(IsDBNull(objRecordSet.Fields("Basic_Amount").Value), 0, objRecordSet.Fields("Basic_Amount").Value)
            If blnISBasicRoundOff = False Then
                dblBasicAmount = System.Math.Round(dblBasicAmount, intBasicRoundOffDecimal)
            ElseIf blnISBasicRoundOff = True Then
                dblBasicAmount = System.Math.Round(dblBasicAmount, 0)
            End If
            If mblnAddCustomerMaterial Then
                dblBaseCurrencyAmount = dblBasicAmount + IIf(IsDBNull(objRecordSet.Fields("CustMtrl_Amount").Value), 0, objRecordSet.Fields("CustMtrl_Amount").Value)
            Else
                dblBaseCurrencyAmount = dblBasicAmount
            End If
            If dblBaseCurrencyAmount > 0 Then
                'initializing the item gl and sl************************
                strRetVal = GetItemGLSL(strGlGroupId, mstrPurposeCode)
                If strRetVal = "N" Then
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetVal, "»")
                strItemGL = varTmp(0)
                strItemSL = varTmp(1)
                'initializing of item gl and sl ends here****************
                rsSalesInvType.ResultSetClose()
                rsSalesInvType = New ClsResultSetDB
                rsSalesInvType.GetResult("Select Invoice_type,Sub_category from SalesChallan_dtl where Unit_code='" & gstrUNITID & "' and Doc_no ='" & strRefInvoiceNo & "'")
                strInvoiceType = rsSalesInvType.GetValue("Invoice_type")
                strInvoiceSubType = rsSalesInvType.GetValue("Sub_category")
                'Posting the basic amount into cost centers, percentage wise
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                objTmpRecordset.Open("SELECT Location_Code,Invoice_Type,Sub_Type,ccM_ccCode,ccM_cc_percentage FROM invcc_dtl WHERE Unit_code='" & gstrUNITID & "' and Invoice_Type='" & strInvoiceType & "' AND Sub_Type = '" & strInvoiceSubType & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' AND ccM_cc_Percentage > 0", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    While Not objTmpRecordset.EOF
                        dblCCShare = (dblBaseCurrencyAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value

                        If gblnGSTUnit = False Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & dblBaseCurrencyAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If

                        objTmpRecordset.MoveNext()
                        iCtr = iCtr + 1
                    End While
                Else
                    If gblnGSTUnit = False Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & dblBaseCurrencyAmount & "»»Basic for Supp Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If

                    iCtr = iCtr + 1
                End If
            End If
            '*********************************************************
            '*********************************************************
            ''EXC Duty Posting
            '*********************************************************
            '*********************************************************
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("TotalExciseAmount").Value), 0, objRecordSet.Fields("TotalExciseAmount").Value)
            If blnISExciseRoundOff = False Then
                dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, intExciseRoundOffDecimal)
            ElseIf blnISExciseRoundOff = True Then
                dblBaseCurrencyAmount = System.Math.Round(dblTaxAmt, 0)
            End If
            If dblBaseCurrencyAmount > 0 Then
                'initializing the tax gl and sl here
                strRetVal = GetTaxGlSl("EXC")
                If strRetVal = "N" Then
                    MsgBox("GL for ARTAX is not defined for EXC", MsgBoxStyle.Information, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetVal, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»EXC»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                iCtr = iCtr + 1
            End If
            objRecordSet.MoveNext()
        End While
        ''Posting of rounded off amount
        strRetVal = GetItemGLSL("", "Rounded_Amt")
        If strRetVal = "N" Then
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        varTmp = Split(strRetVal, "»")
        strItemGL = varTmp(0)
        strItemSL = varTmp(1)
        dblBaseCurrencyAmount = dblInvoiceAmtRoundOff_diff
        dblBaseCurrencyAmount = System.Math.Round(dblBaseCurrencyAmount, intBasicRoundOffDecimal)
        If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»RND»0»" & "»»0»" & strItemGL & "»" & strItemSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»"
            If dblBaseCurrencyAmount < 0 Then
                mstrDetailString = mstrDetailString & "Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "Dr»»»»»»0»0»0»0»0" & "¦"
            End If
        End If
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        CreateStringForAccounts = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        CreateStringForAccounts = False
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetItemGLSL(ByVal InventoryGlGroup As String, ByVal PurposeCode As String) As String
        Dim objRecordSet As New ADODB.Recordset
        Dim strGL As String
        Dim strSL As String
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE Unit_code='" & gstrUNITID & "' and invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE Unit_code='" & gstrUNITID & "' and gbl_prpsCode = '" & PurposeCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRecordSet.EOF Then
                GetItemGLSL = "N"
                MsgBox("GL and SL not defined for Purpose Code: " & PurposeCode, MsgBoxStyle.Information, ResolveResString(100))
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            Else
                strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_glCode").Value), "", objRecordSet.Fields("gbl_glCode").Value))
                strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_slCode").Value), "", objRecordSet.Fields("gbl_slCode").Value))
            End If
        Else
            strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_glcode").Value), "", objRecordSet.Fields("invGld_glcode").Value))
            strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_slcode").Value), "", objRecordSet.Fields("invGld_slcode").Value))
        End If
        If strGL = "" Then
            GetItemGLSL = "N"
            MsgBox("GL and SL not defined for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, ResolveResString(100))
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetItemGLSL = strGL & "»" & strSL
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetItemGLSL = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Private Sub ShowCode_Desc(ByVal pstrQuery As String, ByRef pctlCode As System.Windows.Forms.TextBox, Optional ByRef pctlDesc As System.Windows.Forms.Label = Nothing)
        Dim varHelp As Object
        On Error GoTo ErrHandler
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        'Changing the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, pstrQuery)
        'Changing the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                pctlCode.Text = Trim(varHelp(0))
                If Not (pctlDesc Is Nothing) Then
                    pctlDesc.Text = Trim(varHelp(1))
                End If
                pctlCode.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function GenerateInvoiceNo(ByVal pstrInvoiceType As String, ByRef pstrInvoiceSubType As String, ByVal pstrRequiredDate As String, ByVal strCustomerCode As String) As String
        On Error GoTo ErrHandler
        Dim clsInstEMPDBDbase As New EMPDataBase.EMPDB(gstrUnitId)
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim strSuffix As String 'Generate a NEW Series
        Dim strZeroSuffix As String
        Dim strFin_Start_Date As String
        Dim strFin_End_Date As String
        Dim strsql As String 'String SQL Query
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        If Len(Trim(pstrInvoiceType)) > 0 Then 'For Dated Docs
            strsql = "Set Dateformat 'dmy' Select Current_No,Suffix,Fin_start_date,Fin_end_Date, ISNULL(CURRENT_NO_TRF_SAMEGSTIN,0) CURRENT_NO_TRF From saleConf Where Unit_code='" & gstrUNITID & "' and"
            strsql = strsql & " Invoice_Type ='" & pstrInvoiceType & "' and  sub_type='" & pstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & pstrRequiredDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & pstrRequiredDate & "')<=0"
            With clsInstEMPDBDbase.CConnection
                .OpenConnection(gstrDSNName, gstrDatabaseName)
                .ExecuteSQL("Set Dateformat 'dmy'")
            End With
            clsInstEMPDBDbase.CRecordset.OpenRecordset(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If Not clsInstEMPDBDbase.CRecordset.EOF_Renamed Then
                'Get Last Doc No Saved
                '101188073 Start
                If pstrInvoiceType.ToUpper() = "TRF" Then
                    If IsGSTINSAME(strCustomerCode) Then
                        strCheckDOcNo = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("CURRENT_NO_TRF", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                    Else
                        strCheckDOcNo = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Current_No", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                    End If
                Else
                    strCheckDOcNo = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Current_No", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                End If
                '101188073 End
                strSuffix = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("suffix", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                strFin_Start_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_Start_Date", EMPDataBase.EMPDB.ADODataType.ADODate, EMPDataBase.EMPDB.ADOCustomFormat.CustomDate))
                strFin_End_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_End_Date", EMPDataBase.EMPDB.ADODataType.ADODate, EMPDataBase.EMPDB.ADOCustomFormat.CustomDate))
            Else
                'No Records Found
                Err.Raise(vbObjectError + 20008, "[GenerateDocNo]", "Incorrect Parameters Passed Invoice Number cannot be Generated.")
            End If
            clsInstEMPDBDbase.CRecordset.CloseRecordset() 'Close Recordset
        Else
            'ELSE Raise Error If Wanted Date Not Passed
            Err.Raise(vbObjectError + 20007, "[GenerateDocNo]", "Wanted Date Information not Passed")
        End If
        If Len(Trim(strCheckDOcNo)) > 0 Then 'That is the Document is Made for that Perio
            'Add 1 to it
            strTempSeries = CStr(CInt(strCheckDOcNo) + 1)
            mSaleConfNo = Val(strTempSeries)
            If Len(Trim(strTempSeries)) < 6 Then
                intMaxLoop = 6 - Len(Trim(strTempSeries))
                strZeroSuffix = ""
                For intLoopCounter = 1 To intMaxLoop
                    strZeroSuffix = Trim(strZeroSuffix) & "0"
                Next
            End If
            strTempSeries = strSuffix & strZeroSuffix & strTempSeries
            '101188073 Start
            If gblnGSTUnit Then
                If Len(GSTUnitPrefixCode) > 0 AndAlso GSTUnitPrefixCode <> 0 Then
                    strTempSeries = GSTUnitPrefixCode & strTempSeries
                End If
            End If
            '101188073 End
            'UpDate Back New Number
            GenerateInvoiceNo = strTempSeries
        End If
        Exit Function
ErrHandler:
        'Logging the ERROR at Application's Path
        Dim clsErrorInst As New EMPDataBase.EMPDB(gstrUnitId)
        clsErrorInst.CError.RaiseError(20008, "[frmexptrn0006]", "[GenerateInvoiceNo]", "", "No. Not Generated For DocType = " & pstrInvoiceType & " due to [ " & Err.Description & " ].", My.Application.Info.DirectoryPath, gstrDSNName, gstrDatabaseName)
    End Function
    Private Sub ReplaceJunkCharacters()
        On Error GoTo Errorhandler
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(15), "") 'Remove Uncompress Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(18), "") 'Remove Decompress Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "G", "") 'Remove Bold Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "H", "") 'Remove DeBold Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(12), "") 'Remove DeUnderline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-1", "") 'Remove Underline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-0", "") 'Remove DeUnderline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W1", "") 'Remove DoubleWidth Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W0", "") 'Remove DeDoubleWidth Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "M", "") 'Remove Middle Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "P", "") 'Remove DeMiddle Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "E", "") 'Remove Elite Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "F", "") 'Remove DeElite Character
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function printSuppInvWithoutPrefix(ByVal pstrInvNo As String) As Double
        On Error GoTo Errorhandler
        Dim strsql As String
        Dim strSuffix As String
        Dim rstemp As ClsResultSetDB
        Dim StrInvDate As String
        strsql = "select Invoice_date from supplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        rstemp = New ClsResultSetDB
        rstemp.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        StrInvDate = VB6.Format(rstemp.GetValue("Invoice_Date"), "dd mmm yyyy")
        rstemp.ResultSetClose()
        rstemp = New ClsResultSetDB
        rstemp.GetResult("Select isnull(Suffix,0) as suffix from SaleConf where Unit_code='" & gstrUNITID & "' and Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & StrInvDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & StrInvDate & "')<=0")
        strSuffix = Trim(rstemp.GetValue("Suffix"))
        If rstemp.GetNoRows > 0 Then
            If Val(strSuffix) > 0 Then
                printSuppInvWithoutPrefix = Val(Mid(pstrInvNo, Len(strSuffix) + 1))
            Else
                printSuppInvWithoutPrefix = Val(pstrInvNo)
            End If
        Else
            printSuppInvWithoutPrefix = Val(pstrInvNo)
        End If
        rstemp.ResultSetClose()
        rstemp = Nothing
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub Ctlinvoice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Ctlinvoice.Leave
        Dim clsInvoiceNo As ClsResultSetDB
        Dim strsql As String
        On Error GoTo ErrHandler
        If Not (IsNumeric(Trim(Ctlinvoice.Text))) And Len(Trim(Ctlinvoice.Text)) > 0 Then
            MsgBox("Invoice Number doesn't exist.", MsgBoxStyle.Critical, ResolveResString(100))
            Ctlinvoice.Text = ""
            Ctlinvoice.Focus()
            Exit Sub
        End If
        If Len(Trim(Ctlinvoice.Text)) > 0 Then
            strsql = "Select doc_no,bill_flag from SupplementaryInv_Hdr where Unit_code='" & gstrUNITID & "' and doc_no= '" & Trim(Ctlinvoice.Text) & "' "
            clsInvoiceNo = New ClsResultSetDB
            clsInvoiceNo.GetResult(strsql)
            If clsInvoiceNo.GetNoRows > 0 Then
                ''Generate New Invoice.
                If optInvYes(0).Checked = True Then
                    If clsInvoiceNo.GetValue("Bill_flag") = True Then
                        MsgBox("This Invoice Number is already locked. Please enter any Unlocked Number. ", MsgBoxStyle.Information, ResolveResString(100))
                        Ctlinvoice.Text = ""
                        Ctlinvoice.Focus()
                    End If
                    '' Reprint Invoice.
                Else
                    If clsInvoiceNo.GetValue("Bill_flag") = False Then
                        MsgBox("Enter any Unlocked Challan number.", MsgBoxStyle.Information, ResolveResString(100))
                        Ctlinvoice.Text = ""
                        Ctlinvoice.Focus()
                    End If
                End If
            Else
                MsgBox("Invoice Number doesn't exist.", MsgBoxStyle.Critical, ResolveResString(100))
                Ctlinvoice.Text = ""
                Ctlinvoice.Focus()
            End If
            clsInvoiceNo.ResultSetClose()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '101188073
    Private Function IsGSTINSAME(ByVal strCustomerCode As String) As Boolean
        If strCustomerCode Is Nothing Then Return False
        If Len(strCustomerCode) = 0 Then Return False
        If SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_Id,'') GSTIN_Id From Customer_Mst Where UNIT_CODE='" & gstrUnitId & "' And Customer_Code='" & strCustomerCode & "'") = SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_ID,'') GSTIN_ID From Gen_UnitMaster Where Unt_CodeId='" & gstrUnitId & "'") Then
            Return True
        Else
            Return False
        End If
    End Function
    '101188073
End Class