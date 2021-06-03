Imports System.Windows.Forms
Imports System.Collections
Imports System.Configuration
Imports System.Diagnostics
Imports Infragistics.Win
Imports Infragistics.Win.UltraWinExplorerBar
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinToolbars
Imports DBLotVbnet.common
Imports System.Net.NetworkInformation
Imports System.IO
Imports Infragistics.Win.UltraWinGrid
Imports Microsoft.VisualBasic.FileIO
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS
Imports DBLotVbnet.pln_Module
'Imports CrystalDecisions.CrystalReports.Engine
Public Class MDIMain 
    Dim networkcard() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
    Dim dsUser As DataSet
    Dim A As String
    Dim B As New ReportDocument
    Dim exc As New Microsoft.Office.Interop.Excel.Application
    Dim Clicked As String
    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = workbook.Worksheets
    Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Global.System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticleToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub
    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub
    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub
    Private m_ChildFormNumber As Integer = 0
#Region "UltraExplorerBar1_ItemClick"
    Private Sub UltraExplorerBar1_ItemClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemClick
        'Dim ncQryType As String
        'Dim nvcFieldList As String
        'Dim nvcWhereClause As String
        'Dim nvcVccode As String
        'Dim M02 As DataSet

        'Dim connection As SqlClient.SqlConnection
        'Dim transaction As SqlClient.SqlTransaction
        'Dim transactionCreated As Boolean
        'Dim connectionCreated As Boolean
        'Dim strInvo As String
        'Dim strChqvalue As Double


        'connection = DBEngin.GetConnection(True)
        'connectionCreated = True
        'transaction = connection.BeginTransaction()
        'transactionCreated = True

        Dim i As Integer
        Dim X As Integer
        Dim T01 As DataSet
        Dim DateDiffr As TimeSpan
        Dim _Fromdate As Date
        Dim _Todate As Date

        '  Dim B As New ReportDocument
        '============= This is if u want to close active form ========================
        ' ''For Each ChildForm As Form In Me.MdiChildren
        ' ''    ChildForm.Close()
        ' ''Next
        '==============================================================================
        'VserverTime = Getservertime()

        'CType(Me.UltraToolbarsManager1.Tools("DataBase2"), LabelTool).SharedProps.Caption = "Server Date  : " & FormatDateTime(VserverTime, DateFormat.LongDate) & " : " & FormatDateTime(VserverTime, DateFormat.LongTime)

        '=========================================
        Select Case e.Item.Key

            Case "COM"
                frmCompany_Cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmCompany_Cnt.Show()
            Case "NROT"
                frmSupplier_Cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmSupplier_Cnt.Show()
            Case "NEWL"
                frmLocation_Cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmLocation_Cnt.Show()
            Case "SRF"
                frmSales_Ref_Cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmSales_Ref_Cnt.Show()

            Case "CUS"
                frmCustomer_Cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF".
                frmCustomer_Cnt.Show()
            Case "CUSDIS"
                frmCus_Discount.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmCus_Discount.Show()

            Case "SUP"
                frmSupplier_Cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmSupplier_Cnt.Show()
                'GRECI()
            Case "CNE"
                frmElectrician_cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmElectrician_cnt.Show()
            Case "VM"
                frmVehicle_Cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmVehicle_Cnt.Show()
            Case "PDRC"
                frmCatogery_cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmCatogery_cnt.Show()
            Case "NRC"
                frmRow_Category_cntvb.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmRow_Category_cntvb.Show()
            Case "ITM"
                frmItem_cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmItem_cnt.Show()
            Case "BOM"
                frmBom_Creation.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmBom_Creation.Show()
            Case "KS"
                frmStockBalance_Product.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmStockBalance_Product.Show()
            Case "MOB"
                frmStockBalance_Mobile.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmStockBalance_Mobile.Show()
                'RPTTR()
            Case "WEA"
                frmGRN_uniq.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmGRN_uniq.Show()

            Case "POR"
                frmrptPO.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptPO.Show()
            Case "CRII"
                frmMK_Return_cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmMK_Return_cnt.Show()
            Case "MR"
                frmMK_Return.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmMK_Return.Show()

            Case "SRN"
                frmWastage_cnt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmWastage_cnt.Show()
            Case "WE"
                frmWastage.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmWastage.Show()

            Case "MRR"
                frmrptMKReturn.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptMKReturn.Show()

            Case "STKR"
                frmrptStock_uniq.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptStock_uniq.Show()

            Case "PO"
                frmPO.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmPO.Show()
            Case "CRSST"
                frmGrn_T_Acc.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmGrn_T_Acc.Show()

            Case "SRR"
                frmrptSales_uniq.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptSales_uniq.Show()
            Case "KS"
                frmGood_Transfer.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmGood_Transfer.Show()
            Case "PURR"
                frmOutstanding_Collection.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmOutstanding_Collection.Show()

            Case "PI"
                frmItem_Master.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmItem_Master.Show()

            Case "RLCU"
                frmrptCustomer.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptCustomer.Show()
            Case "SMRG"
                frmSub_Manifacter.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmSub_Manifacter.Show()
            Case "STR"
                frmDirect_Sales.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDirect_Sales.Show()
            Case "MCBA"
                frmLocation.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmLocation.Show()
            Case "SIR"
                frmPacking_Box.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmPacking_Box.Show()
            Case "PRD"
                frmPacking.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmPacking.Show()
            Case "RPP"
                frmrptPacking.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptPacking.Show()
            Case "PUPK"
                frmrptUnpacking.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptUnpacking.Show()
            Case "MPR"
                frmAvg_Price.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmAvg_Price.Show()
            Case "RLSU"
                frmJob_Card_Uniq.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmJob_Card_Uniq.Show()
            Case "RTI"
                frmrptItems.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptItems.Show()

            Case "PSC"
                frmSet.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmSet.Show()
            Case "CUS"
                frmNewCustomer.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmNewCustomer.Show()

            Case "SB"
                frmNewStock_Balance.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmNewStock_Balance.Show()

            Case "RMI"
                frmRow_Material.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmRow_Material.Show()
            Case "RPTP"
                frmItem_Issue_Uniq.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmItem_Issue_Uniq.Show()

            Case "LR"
                frmrptRowMaterial.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmrptRowMaterial.Show()
            Case "MCC"
                frmMain.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmMain.Show()

            Case "SP"
                frmNew_Supplier.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmNew_Supplier.Show()
                strWindowName = "frmNew_Supplier"

            Case "MYM"
                frmUML_Employee.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmUML_Employee.Show()

            Case "KS"
                frmKNT_Spec.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmKNT_Spec.Show()


            Case "WOTDA"
                frmOTD_Analysis.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmOTD_Analysis.Show()

            Case "PRVAC"
                frmProvsActual.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmProvsActual.Show()
                ' CREP()
            Case "FDQ"
                frmForwerd.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmForwerd.Show()
                ' CREP()
            Case "CREP"
                frmSetup_Planner.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmSetup_Planner.Show()
            Case "YRC"
                frmYarn_Request_Conformation.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmYarn_Request_Conformation.Show()

            Case "DELR"
                frmDealy_Reason.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDealy_Reason.Show()
            Case "GCM"
                frmGrige_Confirmation.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmGrige_Confirmation.Show()

            Case "DEFC"
                frmDelivary_Forcust.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDelivary_Forcust.Show()

            Case "OTDMR"
                frmUpdate_Status.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmUpdate_Status.Show()

            Case "PDQ"
                frmTjl_Projection.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmTjl_Projection.Show()

            Case "MOQ"
                frmMOQ.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmMOQ.Show()

            Case "DELN"
                frmDealy.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDealy.Show()
                ' BAKR()

            Case "BAKR"
                frmBackLog.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmBackLog.Show()

            Case "OTD"
                frmOTD.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmOTD.Show()
                'DELCUS()
            Case "DELCUS"
                frmDel_Cus.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDel_Cus.Show()

            Case "DRMER"
                frmDelivary_Revision_Merch.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDelivary_Revision_Merch.Show()
            Case "UPPR"
                frmProjection_Upload.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmProjection_Upload.Show()

                ' SRPCL()
                ' LAPDI()
            Case "LAPDI"
                'frmLabDip.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmLabDip.Show()
                'DREP()
            Case "DREP"
                frmDelivery_Revision_Pln.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDelivery_Revision_Pln.Show()
                ' PORR()
            Case "PORR"
                frmPending_Orders.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmPending_Orders.Show()

            Case "OSR"
                frmOrder_Status_Report.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmOrder_Status_Report.Show()

            Case "SRPCL"
                'frmSample_Payment.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmSample_Payment.Show()


            Case "CHKIN"
                'frmPayment_ReceiveInvoice.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmPayment_ReceiveInvoice.Show()

            Case "BYMF"
                'frmLeger_Acc.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmLeger_Acc.Show()

            Case "SSPY"
                'frmStaff_Sallary.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmStaff_Sallary.Show()


            Case "STF"
                'frmMMaster.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmMMaster.Show()

            Case "LJSR"
                'frmLJSR.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmLJSR.Show()

            Case "QAAR"
               

            Case "EWEF"
                frmExaminnerEff.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmExaminnerEff.Show()

            Case "QLIWR"
                frmQualityWise.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmQualityWise.Show()

            Case "EXDT"
                frmExaminer_Down.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmExaminer_Down.Show()

            Case "MDR"
                'frmDowntimereport.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmDowntimereport.Show()
            Case "DRW"
                'frmRSDown.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmRSDown.Show()

            Case "DTMW"
                'frmMCDown.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmMCDown.Show()
            Case "SKNT"
                frmDowntime.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmDowntime.Show()
            Case "RPTF"
                frmFultRate.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmFultRate.Show()

            Case "CPIGC"
                frmCPIChart.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmCPIChart.Show()

            Case "KRS"
                frmKnittingReportShift.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmKnittingReportShift.Show()
            Case "PRRE"
                frmProduction_Examinner.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmProduction_Examinner.Show()
            Case "PRMW"
                frmProduction_MC.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmProduction_MC.Show()


            Case "SD"
                'frmSD.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmSD.Show()
            Case "TCR"
                frmTeco.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmTeco.Show()
            Case "KPR"
                frmKnittingReport.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmKnittingReport.Show()
            Case "PBP"
                'frmReceive_Payment.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmReceive_Payment.Show()
            Case "CPR"
                'frmPI_Confirmation.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmPI_Confirmation.Show()

            Case "DBP"
                'frmInvoiceLJ.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmInvoiceLJ.Show()

            Case "TCP"
                frmQuality.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmQuality.Show()
            Case "RPR"
                frmReprocess.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmReprocess.Show()
            Case "PDU"
                frmProductionReport.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmProductionReport.Show()
            Case "ARR"
                AuditReport.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                AuditReport.Show()

            Case "FBR"
                frmFeedback.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmFeedback.Show()

            Case "FLBU"
                frmRollwt.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmRollwt.Show()

            Case "KPRT"
                frmKnittingProduction.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmKnittingProduction.Show()

            Case "FLBQ"
                frmQReport.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmQReport.Show()
            Case "FLBS"
                frmCPI.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmCPI.Show()
            Case "RFLD"
                frmCutoff.MdiParent = Me
                m_ChildFormNumber += 1
                frmCutoff.Show()

            Case "RFLA"
                frmInspect.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmInspect.Show()

            Case "FLR"
                'frmPI.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmPI.Show()

            Case "SCPI"
                frmSetupCPI.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmWinner.Text = "WNF"
                frmSetupCPI.Show()
            Case "FL"
                'frmFLRequest.MdiParent = Me
                'm_ChildFormNumber += 1
                '' frmWinner.Text = "WNF"
                'frmFLRequest.Show()
            Case "SSN"
                'frmSearchSerial.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmSearchSerial.Show()

            Case "ITEM"
                'frmSalesDetail.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmSalesDetail.Show()
            Case "RPUR"
                'frmSupplier_Master.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmSupplier_Master.Show()

            Case "WQRF"
                frmQualityReport.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmQualityReport.Show()

            Case "GRN"
                ' If strUserLevel = "5" Then
                'frmByuer.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmByuer.Show()

            Case "LEA"
                ' If strUserLevel = "5" Then
                'frmLeave_Form.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmLeave_Form.Show()

                ''End If
            Case "RSM"

                'frmEMP.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmEMP.Show()


            Case "RSSP"
                'frmCustomer.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmCustomer.Show()


            Case "SPR"
                'frmKpro.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmKpro.Show()

            Case "PBBQ"
                frmDNHReport.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmDNHReport.Show()


            Case "QG"
                frmQurantineGrf.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmQurantineGrf.Show()

            Case "QP"
                'frmGRN.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmGRN.Show()
            Case "CFN"
                'frmCancelKnitting.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmCancelKnitting.Show()

            Case "QTR"
                frmQurantine.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmQurantine.Show()
            Case "PBP1"
                frmReport_Barcode.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmReport_Barcode.Show()

            Case "PMR"
                frmProduction.MdiParent = Me
                m_ChildFormNumber += 1
                ' frmVoucher_Detailes.Text = "VOD"
                frmProduction.Show()
            Case "DOWN"
                frmKnittingMC.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmKnittingMC.Show()

            Case "TGS"
                frmTypeofDye.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmTypeofDye.Show()
            Case "KDL"
                'frmKnitting.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmKnitting.Show()
            Case "DOM"
                frmUpload.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmUpload.Show()
                ' ERUDL()

            Case "ERUDL"
                frmEditERAL.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmEditERAL.Show()

            Case "MWDT"
                frmMachine_Downtime.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmMachine_Downtime.Show()

            Case "FIR"
                frmOEE.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmOEE.Show()


            Case "MB51"
                frmUploadMB51.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmUploadMB51.Show()
            Case "DCAA"
                frmAlarm.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmAlarm.Show()
            Case "MRS"
                frmUploadMRS.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmUploadMRS.Show()

            Case "ZDCA"
                frmUpload_ZDCA_PURCHASE.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmUpload_ZDCA_PURCHASE.Show()

            Case "DCA"
                frmDCA_CONSUMPTION.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmDCA_CONSUMPTION.Show()

            Case "FDP"
                frmFDP.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmFDP.Show()

            Case "MTRGR"
                frmMaterial_Group.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmMaterial_Group.Show()



            Case "SAPP"
                frmUploadPlaning.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmUploadPlaning.Show()
            Case "LIB"
                frmLIB.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmLIB.Show()

            Case "DAYPL"
                frmUpload_Dayplane.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmUpload_Dayplane.Show()

            Case "BLM"
                frmBlockMT.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmBlockMT.Show()


            Case "PLCOM"
                frmPlnComm.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmPlnComm.Show()

                'SEASDE()
            Case "SEASDE"
                'frmBOffice.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmBOffice.Show()

            Case "NBO"
                'frmBOffice.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmBOffice.Show()

            Case "SAMP"
                frmSample.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmSample.Show()

            Case "SDTIM"
                frmSetupTime.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmSetupTime.Show()
                ' PRFGU()
            Case "PRFGU"
                frmDNFProduction.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmDNFProduction.Show()
            Case "AC"
                frmAccount_Statment.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmAccount_Statment.Show()
            Case "PU"
                frmMakeaPayment.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmMakeaPayment.Show()

            Case "DUTC"
                frmEngineering.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmEngineering.Show()


            Case "PIF"
                frmPigment.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmPigment.Show()
            Case "PI"
                frmPr_Item.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmPr_Item.Show()

            Case "MRSREP"
                MRSReport.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                MRSReport.Show()


            Case "ALARMREP"
                frmAlarmReport.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmAlarmReport.Show()

                ' DCAOA()

            Case "DCAOA"
                frmOrderAnalysis.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmOrderAnalysis.Show()

            Case "OPRPO"
                frmPurchasing_Report.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmPurchasing_Report.Show()
            Case "PLNRT"
               

            Case "DEIR"
                frmDelivery.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmDelivery.Show()
            Case "LBM"
                'frmFactory.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmFactory.Show()
            Case "LBDA"
                'frmLab_DipApp.MdiParent = Me
                'm_ChildFormNumber += 1
                ''frmVoucher_Detailes.Text = "VOD"
                'frmLab_DipApp.Show()
            Case "PRINTDEL"
                frmApp_DR.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmApp_DR.Show()
            Case "NOOR"
                frmNo_Orders.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmNo_Orders.Show()

            Case "FGSA"
                frmFG_Stock.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmFG_Stock.Show()

            Case "GPRP"
                frmGrgProvision.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmGrgProvision.Show()

            Case "CRMM"
                frmMerchantMaster.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmMerchantMaster.Show()


            Case "OPRPT"
                frmOrder_Placement.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmOrder_Placement.Show()
            Case "TTTFG"
                frmTime_Taken.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmTime_Taken.Show()

            Case "CCCS"
                frmNewCustomer.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmNewCustomer.Show()

            Case "VIRR"
                frmVarions_rpt.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmVarions_rpt.Show()
            Case "SSCR"
                frmCr_Received.MdiParent = Me
                m_ChildFormNumber += 1
                'frmVoucher_Detailes.Text = "VOD"
                frmCr_Received.Show()
        End Select

    End Sub




#End Region
#Region "MDIMain_Load"

    Private Sub MDIMain_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
        SPL.Close()
        frmViewPO.Close()
        frmViewGRN.Close()
        Me.Close()
    End Sub
    'Function Load_Img()
    '    Dim _Date As Date
    '    Dim i As Integer
    '    Dim _Month As Integer
    '    Dim strFilePath As String

    '    _Date = Today
    '    For i = 1 To 6
    '        _Month = Month(_Date)

    '        'If Microsoft.VisualBasic.Day(_Date) <= 10 Then
    '        '    _Date = _Date.AddDays(+30)
    '        'ElseIf Microsoft.VisualBasic.Day(_Date) >= 20 Then
    '        '    _Date = _Date.AddDays(+15)
    '        'End If
    '        If i = 1 Then
    '            strFilePath = ConfigurationManager.AppSettings("imgPath") + "\Icons8-Windows-8-Animals-Gorilla.ico"

    '            UltraToolbarsManager1.Ribbon.Tabs(0).Groups(4).Tools(0).SharedProps.Caption = MonthName(_Month)
    '            UltraToolbarsManager1.Ribbon.Tabs(0).Groups(4).Tools(0).SharedProps.AppearancesSmall.Appearance.Image = strFilePath
    '            UltraToolbarsManager1.Ribbon.Tabs(0).Groups(4).Tools(0).Reset()
    '            '  UltraToolbarsManager1.Ribbon.Tabs(0).Groups(4).Tools(0).SharedProps.Dispose()
    '        End If
    '        If Microsoft.VisualBasic.Day(_Date) <= 10 Then
    '            _Date = _Date.AddDays(+30)
    '        ElseIf Microsoft.VisualBasic.Day(_Date) >= 20 Then
    '            _Date = _Date.AddDays(+15)
    '        End If
    '    Next
    'End Function

    Private Sub MDIMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim SQL As String
        'Dim con = New SqlConnection()

        Try

            ' con = DBEngin.GetConnection()
            'Infragistics.Win.AppStyling.StyleManager.Load(ConfigurationManager.AppSettings("APPSTL"))
            Infragistics.Win.AppStyling.StyleManager.Load(ConfigurationManager.AppSettings("APPSTL"))
            Nav_Load()
            ' VserverTime = Getservertime()
            'TodayDraw()
            'Try
            'CType(Me.UltraToolbarsManager1.Tools("MName"), LabelTool).SharedProps.Caption = "Machine Name : " & (VMName) & " (" & netCard & ")"
            'CType(Me.UltraToolbarsManager1.Tools("MIP"), LabelTool).SharedProps.Caption = "Machine IP : " & VIP
            'CType(Me.UltraToolbarsManager1.Tools("MUser"), LabelTool).SharedProps.Caption = "Logged User : " & LoggedUser
            'CType(Me.UltraToolbarsManager1.Tools("DataSource2"), LabelTool).SharedProps.Caption = "Data Source : " & VDataSource & " : " & VDataBase
            'CType(Me.UltraToolbarsManager1.Tools("DataBase2"), LabelTool).SharedProps.Caption = "Server Date  : " & FormatDateTime(VserverTime, DateFormat.LongDate) & " : " & FormatDateTime(VserverTime, DateFormat.LongTime)

            Me.UltraToolbarsManager1.Ribbon.Tabs(0).Groups(2).Visible = False


            ' Catch
            ' MsgBox("Unable to Return your Network Details", MsgBoxStyle.Critical)
            'End Try
            '======================= Load C01Workstation Data to Main Menu
            ' Workstation_Load()
            '===========================================================
            'UltraExplorerBar1.Groups(0).Visible = False
            UltraExplorerBar1.Groups(1).Visible = False
            UltraExplorerBar1.Groups(2).Visible = False
            UltraExplorerBar1.Groups(3).Visible = False
            UltraExplorerBar1.Groups(4).Visible = False
            ' UltraExplorerBar1.Groups(5).Visible = False
            'UltraExplorerBar1.Groups(6).Visible = False
            'UltraExplorerBar1.Groups(7).Visible = False

            SPL.Hide()
            frmLog.MdiParent = Me
            frmLog.Show()

            strKey = netCard
            '  Call frmCost.Load_Amount()
            ' Call Load_Img()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If

        End Try
    End Sub
#End Region

#Region "Ribbon Click"
    Private Sub UltraToolbarsManager1_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        Select Case e.Tool.Key
            Case "cmb_mdi"
                'If Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "WNF" Then
                '    frmWinner.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmWinner.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "AGT" Then
                '    frmAgent_Maintenance.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmAgent_Maintenance.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "PPC" Then
                '    frmPayment_Cancel.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmPayment_Cancel.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "PDV" Then
                '    frmPayment_Duplicate.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmPayment_Duplicate.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "MCP" Then
                '    frmMiscellaneous.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmMiscellaneous.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "VOD" Then
                '    frmVoucher_Detailes.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmVoucher_Detailes.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "IST" Then
                '    frmInterLocation.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmInterLocation.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "ITD" Then
                '    frmVoucher_Detailes.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmVoucher_Detailes.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "IRD" Then
                '    frmVoucher_Detailes.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmVoucher_Detailes.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "ISP" Then
                '    frmVoucher_Detailes.MdiParent = Me
                '    m_ChildFormNumber += 1
                '    frmVoucher_Detailes.Show()
                'ElseIf Mid(CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).Text, 1, 3) = "XXX" Then
                '    For Each ChildForm As Form In Me.MdiChildren
                '        ChildForm.Close()
                '    Next
                '    Me.Close()
                'End If
        End Select
    End Sub
#End Region
    Sub Nav_Load()
        'Dim i As Integer
        'Dim con = New SqlConnection()
        'Dim valueList As ValueList = New ValueList()

        ''=======================================================================
        'Try
        '    con = DBEngin.GetConnection()
        '    'dsUser = DBEngin.ExecuteDataset(con, Nothing, "Lst_Navigation", New SqlParameter("@cQryType", "NAV"))
        '    'For i = 0 To dsUser.Tables(0).Rows.Count - 1
        '    '    valueList.ValueListItems.Add(i, (dsUser.Tables(0).Rows(i)("SHT_Key") & " - " & (dsUser.Tables(0).Rows(i)("SHT_Description"))))
        '    'Next i
        '    'CType(Me.UltraToolbarsManager1.Tools("cmb_mdi"), ComboBoxTool).ValueList = valueList

        '    '===================================================================
        'Catch returnMessage As Exception
        '    If returnMessage.Message <> Nothing Then
        '        MessageBox.Show(returnMessage.Message)
        '    End If
        'End Try
        ''=========================================================================
        ' "asdasd"
    End Sub
#Region "Workstation_Load"
    Sub Workstation_Load()
        Dim i As Integer
        ' Dim con = New SqlConnection()

        '=======================================================================
        Try
            '    con = DBEngin.GetConnection(True)
            'dsUser = DBEngin.ExecuteDataset(con, Nothing, "Lst_Navigation", New SqlParameter("@cQryType", "MAC"), New SqlParameter("@cMacCode", netCard))

            'For i = 0 To dsUser.Tables(0).Rows.Count - 1

            '    If dsUser.Tables(0).Rows(i)("c01systemstart") = "C" Then

            ' UltraExplorerBar1.Groups(0).Items(3).Visible = False
            '  UltraExplorerBar1.Groups(1).Items(4).Visible = False

            'UltraExplorerBar1.Groups(1).Items(4).Visible = False

            If Trim(strUGroup) = "PLN" Then
                UltraExplorerBar1.Groups(0).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(2).Items(0).Visible = False
                UltraExplorerBar1.Groups(2).Items(1).Visible = True
                UltraExplorerBar1.Groups(2).Items(4).Visible = True

           
                '   UltraExplorerBar1.Groups(4).Items(1).Visible = False
            ElseIf Trim(strUGroup) = "ADMIN" Then
                UltraExplorerBar1.Groups(0).Visible = True
                '  UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                'UltraExplorerBar1.Groups(4).Visible = True
                UltraExplorerBar1.Groups(3).Visible = True
                UltraExplorerBar1.Groups(5).Visible = True
                ' UltraExplorerBar1.Groups(7).Visible = True
                UltraExplorerBar1.Groups(2).Items(0).Visible = True
                ' UltraExplorerBar1.Groups(2).Items(1).Visible = True
                UltraExplorerBar1.Groups(2).Items(2).Visible = True
                'UltraExplorerBar1.Groups(2).Items(3).Visible = True
                UltraExplorerBar1.Groups(2).Items(6).Visible = True
                '  UltraExplorerBar1.Groups(2).Items(7).Visible = True
            ElseIf Trim(strUGroup) = "SUPPER" Then
                UltraExplorerBar1.Groups(0).Visible = True
                UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(4).Visible = True
                UltraExplorerBar1.Groups(3).Visible = True
                UltraExplorerBar1.Groups(2).Items(0).Visible = True
                UltraExplorerBar1.Groups(2).Items(1).Visible = False
                UltraExplorerBar1.Groups(2).Items(2).Visible = True
                UltraExplorerBar1.Groups(2).Items(3).Visible = False
                UltraExplorerBar1.Groups(3).Items(2).Visible = False

            ElseIf Trim(strUGroup) = "STORS" Then
                UltraExplorerBar1.Groups(0).Visible = True
                UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(4).Visible = True
                UltraExplorerBar1.Groups(3).Visible = True
                UltraExplorerBar1.Groups(2).Items(0).Visible = True
                UltraExplorerBar1.Groups(2).Items(1).Visible = False
                UltraExplorerBar1.Groups(2).Items(2).Visible = True
                UltraExplorerBar1.Groups(2).Items(3).Visible = True
                UltraExplorerBar1.Groups(3).Items(2).Visible = False
                UltraExplorerBar1.Groups(2).Items(6).Visible = False

            ElseIf Trim(strUGroup) = "PROCU" Then
                UltraExplorerBar1.Groups(0).Visible = False
                UltraExplorerBar1.Groups(1).Visible = False
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(3).Visible = False
                UltraExplorerBar1.Groups(4).Visible = False
                UltraExplorerBar1.Groups(2).Items(0).Visible = False
                UltraExplorerBar1.Groups(2).Items(1).Visible = False
                UltraExplorerBar1.Groups(2).Items(2).Visible = False
                UltraExplorerBar1.Groups(2).Items(3).Visible = False
                UltraExplorerBar1.Groups(2).Items(4).Visible = False
                UltraExplorerBar1.Groups(2).Items(5).Visible = False
                UltraExplorerBar1.Groups(2).Items(6).Visible = False
                UltraExplorerBar1.Groups(2).Items(7).Visible = False
                UltraExplorerBar1.Groups(2).Items(9).Visible = True
                UltraExplorerBar1.Groups(2).Items(10).Visible = False
                UltraExplorerBar1.Groups(2).Items(11).Visible = False
                UltraExplorerBar1.Groups(2).Items(8).Visible = False

                ' UltraExplorerBar1.Groups(0).Items(1).Visible = False

                UltraExplorerBar1.Groups(4).Visible = False
            ElseIf Trim(strUserLevel) = "09" Then
                UltraExplorerBar1.Groups(4).Visible = True
                UltraExplorerBar1.Groups(1).Visible = False
                UltraExplorerBar1.Groups(2).Visible = False
                UltraExplorerBar1.Groups(0).Visible = False
                UltraExplorerBar1.Groups(4).Items(0).Visible = False
                UltraExplorerBar1.Groups(4).Items(1).Visible = True
                UltraExplorerBar1.Groups(4).Items(2).Visible = True
                UltraExplorerBar1.Groups(4).Items(3).Visible = True
                UltraExplorerBar1.Groups(4).Items(5).Visible = False
                UltraExplorerBar1.Groups(4).Items(6).Visible = False
                UltraExplorerBar1.Groups(4).Items(7).Visible = False

                UltraExplorerBar1.Groups(3).Visible = False
            ElseIf Trim(strUserLevel) = "04" Then
                UltraExplorerBar1.Groups(0).Items(2).Visible = True
                UltraExplorerBar1.Groups(1).Items(4).Visible = True
                UltraExplorerBar1.Groups(0).Visible = True
                UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(4).Visible = True
                UltraExplorerBar1.Groups(3).Visible = True
                '  UltraExplorerBar1.Groups(3).Visible = True
            ElseIf Trim(strUserLevel) = "02" Then
                UltraExplorerBar1.Groups(0).Items(2).Visible = True
                UltraExplorerBar1.Groups(0).Visible = True
                UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(4).Visible = False

            ElseIf Trim(strUserLevel) = "03" Then
                '  UltraExplorerBar1.Groups(1).Items(4).Visible = False
                UltraExplorerBar1.Groups(1).Items(4).Visible = True
                UltraExplorerBar1.Groups(0).Visible = True
                UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(3).Visible = True
                UltraExplorerBar1.Groups(4).Visible = False
            ElseIf Trim(strUserLevel) = "05" Then
                UltraExplorerBar1.Groups(1).Items(0).Visible = False
                UltraExplorerBar1.Groups(1).Items(2).Visible = False
                UltraExplorerBar1.Groups(1).Items(1).Visible = False
                UltraExplorerBar1.Groups(1).Items(3).Visible = False
                UltraExplorerBar1.Groups(1).Items(4).Visible = False
                UltraExplorerBar1.Groups(1).Items(5).Visible = True
                UltraExplorerBar1.Groups(1).Items(6).Visible = False
                UltraExplorerBar1.Groups(1).Items(7).Visible = False
                UltraExplorerBar1.Groups(1).Items(8).Visible = False


                UltraExplorerBar1.Groups(0).Visible = False
                UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = False
                UltraExplorerBar1.Groups(3).Visible = True
                UltraExplorerBar1.Groups(4).Visible = True
                ' UltraExplorerBar1.Groups(3).Visible = True

                UltraExplorerBar1.Groups(3).Items(2).Visible = False
                UltraExplorerBar1.Groups(3).Items(3).Visible = False
                UltraExplorerBar1.Groups(3).Items(4).Visible = False
                UltraExplorerBar1.Groups(3).Items(6).Visible = False
                UltraExplorerBar1.Groups(3).Items(7).Visible = False
                UltraExplorerBar1.Groups(3).Items(8).Visible = False
                UltraExplorerBar1.Groups(3).Items(9).Visible = False
                ' UltraExplorerBar1.Groups(2).Items(10).Visible = False
                UltraExplorerBar1.Groups(3).Items(11).Visible = False
                UltraExplorerBar1.Groups(3).Items(12).Visible = False
                UltraExplorerBar1.Groups(3).Items(13).Visible = False
                UltraExplorerBar1.Groups(3).Items(14).Visible = False
                '  UltraExplorerBar1.Groups(2).Items(16).Visible = False
                UltraExplorerBar1.Groups(3).Items(18).Visible = True
                UltraExplorerBar1.Groups(3).Items(19).Visible = False
                UltraExplorerBar1.Groups(3).Items(20).Visible = False
                UltraExplorerBar1.Groups(3).Items(21).Visible = False
                UltraExplorerBar1.Groups(3).Items(22).Visible = False


            ElseIf Trim(strUserLevel) = "06" Then
                '  UltraExplorerBar1.Groups(0).Items(2).Visible = True
                UltraExplorerBar1.Groups(1).Items(4).Visible = True
                UltraExplorerBar1.Groups(1).Items(0).Visible = False
                UltraExplorerBar1.Groups(1).Items(5).Visible = False
                UltraExplorerBar1.Groups(1).Items(2).Visible = True
                UltraExplorerBar1.Groups(1).Items(3).Visible = True

                UltraExplorerBar1.Groups(0).Visible = False
                UltraExplorerBar1.Groups(1).Visible = True
                UltraExplorerBar1.Groups(2).Visible = True
                UltraExplorerBar1.Groups(3).Visible = True
                UltraExplorerBar1.Groups(4).Visible = False

            ElseIf Trim(strUserLevel) = "07" Then


                UltraExplorerBar1.Groups(0).Visible = False
                UltraExplorerBar1.Groups(1).Visible = False
                UltraExplorerBar1.Groups(2).Visible = False
                UltraExplorerBar1.Groups(3).Visible = False
                UltraExplorerBar1.Groups(4).Visible = True
                UltraExplorerBar1.Groups(4).Items(1).Visible = False
                UltraExplorerBar1.Groups(4).Items(1).Visible = False
                UltraExplorerBar1.Groups(4).Items(2).Visible = False
                UltraExplorerBar1.Groups(4).Items(3).Visible = False
                UltraExplorerBar1.Groups(4).Items(4).Visible = False
            Else
                'UltraExplorerBar1.Groups(0).Visible = False
                'UltraExplorerBar1.Groups(1).Visible = True
                'UltraExplorerBar1.Groups(2).Visible = False
                'UltraExplorerBar1.Groups(4).Visible = False

            End If
            'UltraExplorerBar1.Groups(1).Visible = True
            'UltraExplorerBar1.Groups(2).Visible = True
            'UltraExplorerBar1.Groups(3).Visible = True
            '  UltraExplorerBar1.Groups(4).Visible = True
            ' UltraExplorerBar1.Groups(5).Visible = True
            'UltraExplorerBar1.Groups(6).Visible = True
            'UltraExplorerBar1.Groups(7).Visible = True
            'UltraExplorerBar1.Groups(8).Visible = True
            'UltraExplorerBar1.Groups(9).Visible = True
            '    ElseIf dsUser.Tables(0).Rows(i)("c01systemstart") = "B" Then
            'UltraExplorerBar1.Groups(3).Visible = True
            'UltraExplorerBar1.Groups(4).Visible = False

            'UltraExplorerBar1.Groups(1).Enabled = False
            'UltraExplorerBar1.Groups(3).Enabled = True
            'UltraExplorerBar1.Groups(5).Enabled = False
            'UltraExplorerBar1.Groups(0).Enabled = False

            '    ElseIf dsUser.Tables(0).Rows(i)("c01systemstart") = "S" Then
            'UltraExplorerBar1.Groups(3).Visible = True
            'UltraExplorerBar1.Groups(4).Visible = True

            'UltraExplorerBar1.Groups(1).Enabled = False
            'UltraExplorerBar1.Groups(2).Enabled = False
            'UltraExplorerBar1.Groups(5).Enabled = False
            'UltraExplorerBar1.Groups(0).Enabled = False

            '    ElseIf dsUser.Tables(0).Rows(i)("c01systemstart") = "D" Then
            'UltraExplorerBar1.Groups(4).Visible = True
            'UltraExplorerBar1.Groups(5).Visible = True

            'UltraExplorerBar1.Groups(1).Enabled = False
            'UltraExplorerBar1.Groups(2).Enabled = False
            'UltraExplorerBar1.Groups(3).Enabled = False
            'UltraExplorerBar1.Groups(0).Enabled = False

            '    ElseIf dsUser.Tables(0).Rows(i)("c01systemstart") = "P" Then
            'UltraExplorerBar1.Groups(0).Visible = True
            'UltraExplorerBar1.Groups(4).Visible = True

            'UltraExplorerBar1.Groups(1).Enabled = False
            'UltraExplorerBar1.Groups(2).Enabled = False
            'UltraExplorerBar1.Groups(3).Enabled = False
            'UltraExplorerBar1.Groups(5).Enabled = False


            '    ElseIf dsUser.Tables(0).Rows(i)("c01systemstart") = "R" Then
            'UltraExplorerBar1.Groups(1).Visible = True
            'UltraExplorerBar1.Groups(4).Visible = True

            'UltraExplorerBar1.Groups(0).Enabled = False
            'UltraExplorerBar1.Groups(2).Enabled = False
            'UltraExplorerBar1.Groups(3).Enabled = False
            'UltraExplorerBar1.Groups(5).Enabled = False

            '===================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
        '=========================================================================

    End Sub
#End Region


#Region " UpdateStatusBar "

    Private Sub UpdateStatusBar(ByVal caption As String)

        ' strip ampersand
        Dim index As Integer = caption.IndexOf("&"c)
        If index <> -1 Then
            caption = caption.Remove(index, 1)

        End If

        Me.UltraStatusBar1.Panels("currentToolPanel").Text = caption

    End Sub 'UpdateStatusBar

#End Region


   
End Class
