Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
'Imports CrystalDecisions.CrystalReports.Engine
Public Class frmSpairpartrpt
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim dblInsuaranceCommision As Double
    Dim c_dataCustomer As DataTable
    Dim strPrice As Double
    Dim strTicket_price As Double
    Dim Cmax As String
    Dim strNetamount As Double
    Dim strDiscount As Double
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        'OPR3.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        cmdSave.Enabled = True
        txtDate.Text = Today
        txtTo.Text = Today
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        OPR0.Enabled = False
        'OPR3.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        'Dim B As New ReportDocument
        'Dim A As String
        'Dim StrFromDate As String
        'Dim StrToDate As String
        'Try
        '    A = ConfigurationManager.AppSettings("ReportPath") + "\Spirptrpt.rpt"


        '    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
        '    StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

        '    B.Load(A.ToString)
        '    B.SetParameterValue("Todate", txtTo.Value)
        '    B.SetParameterValue("Fromdate", txtDate.Value)
        '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
        '    frmReport.CrystalReportViewer1.DisplayToolbar = True
        '    frmReport.CrystalReportViewer1.SelectionFormula = "{T04ItemSalesHeader.T04Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
        '    frmReport.MdiParent = MDIMain
        '    'frmReport.Show()
        '    frmReport.Show()
        Heading()
        'Catch returnMessage As Exception
        '    If returnMessage.Message <> Nothing Then
        '        MessageBox.Show(returnMessage.Message)
        '    End If
        'End Try
    End Sub

    Function Heading()
        Dim cSqlStr As String
        Dim Sql As String
        Dim con = New SqlConnection()
        Dim i As Integer
        Dim x As Integer
        Dim recT03 As DataSet
        Dim Y As Integer
        con = DBEngin.GetConnection()
        Try
            Sql = "select * from T05SalesHeader inner join T06SalesFluter on T06InvoiceNo=T05Invoice inner join M04Item on T06ItemCode=M04ItemCode where T05Status='A' and T05Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and M04Status='A'"
            recT03 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            Y = 0
            Dim nPageNo, nPageLen As Integer
            Dim nUserTotal As Double
            Dim cUser As String
            Dim strPanaltydate As Integer
            Dim strTotal As Double
            Dim strAdvance As Double
            nPageLen = 72
            'lblPro.Text = "Connecting .... "
            'lblPro.Refresh()

            i = 0
            'pbCount.Minimum = 0
            'pbCount.Value = pbCount.Minimum
            'pbCount.Maximum = recT03.Tables(0).Rows.Count

            FileOpen(1, "Lmxprint.rpt", OpenMode.Output)
            PrintLine(1, Chr(27) & Chr(64))
            PrintLine(1, Chr(27) & "CH")
            PageHeader()
            nUserTotal = 0
            strTotal = 0
            strAdvance = 0
            For Each DTRow As DataRow In recT03.Tables(0).Rows
                PrintLine(1, TAB(2), recT03.Tables(0).Rows(i)("T05Invoice"), TAB(18), recT03.Tables(0).Rows(i)("M04ItemName"), TAB(42), recT03.Tables(0).Rows(i)("T06Qty"), TAB(59), VB6.Format(recT03.Tables(0).Rows(i)("T06Rate"), "#.00"), TAB(72), VB6.Format(recT03.Tables(0).Rows(i)("T06Discount"), "#.00"), TAB(86), ".00", TAB(97), VB6.Format(recT03.Tables(0).Rows(i)("T06Total"), "#.00"))
                'nUserTotal = nUserTotal + dsUser.Tables(0).Rows(Y)("T04Amount")
                'strPanaltydate = CDate(txtDate.Text).ToOADate - CDate(recT03.Tables(0).Rows(i)("T02NextPaidday")).ToOADate
                strTotal = strTotal + Val(recT03.Tables(0).Rows(i)("T06Total"))
                nUserTotal = nUserTotal + Val(recT03.Tables(0).Rows(i)("T06Discount"))

                ' pbCount.Value = pbCount.Value + 1
                i = i + 1
            Next
            i = 0
            Sql = "select * from T05SalesHeader inner join T06SalesFluter on T06InvoiceNo=T05Invoice inner join M04Item on T06ItemCode=M04ItemCode where T05Status='S' or T05Status='C' and T05Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and M04Status='A'"
            recT03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If recT03.Tables(0).Rows.Count > 0 Then
                PrintLine(1, "")
                PrintLine(1, "Advance for Jewelary")
                PrintLine(1, "---------------------")
            End If
            For Each DTRow As DataRow In recT03.Tables(0).Rows
                PrintLine(1, TAB(2), recT03.Tables(0).Rows(i)("T05Invoice"), TAB(18), recT03.Tables(0).Rows(i)("M04ItemName"), TAB(42), recT03.Tables(0).Rows(i)("T06Qty"), TAB(59), VB6.Format(recT03.Tables(0).Rows(i)("T06Rate"), "#.00"), TAB(72), ".00", TAB(86), VB6.Format(recT03.Tables(0).Rows(i)("T06Discount"), "#.00"), TAB(97), VB6.Format(recT03.Tables(0).Rows(i)("T06discount"), "#.00"))
                'nUserTotal = nUserTotal + dsUser.Tables(0).Rows(Y)("T04Amount")
                'strPanaltydate = CDate(txtDate.Text).ToOADate - CDate(recT03.Tables(0).Rows(i)("T02NextPaidday")).ToOADate
                strTotal = strTotal + Val(recT03.Tables(0).Rows(i)("T06Discount"))
                strAdvance = strAdvance + Val(recT03.Tables(0).Rows(i)("T06Discount"))

                ' pbCount.Value = pbCount.Value + 1
                i = i + 1
            Next
            i = 0
            Sql = "select * from T05SalesHeader inner join T07AdvanceSales on T07InvoiceNo=T05Invoice  where T05Status='S' or T05Status='C' and T05Date between '" & txtDate.Text & "' and '" & txtTo.Text & "'"
            recT03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow As DataRow In recT03.Tables(0).Rows
                PrintLine(1, TAB(2), recT03.Tables(0).Rows(i)("T05Invoice"), TAB(18), "", TAB(42), "", TAB(59), "", TAB(72), ".00", TAB(86), VB6.Format(recT03.Tables(0).Rows(i)("T07Amount"), "#.00"), TAB(97), VB6.Format(recT03.Tables(0).Rows(i)("T07Amount"), "#.00"))
                'nUserTotal = nUserTotal + dsUser.Tables(0).Rows(Y)("T04Amount")
                'strPanaltydate = CDate(txtDate.Text).ToOADate - CDate(recT03.Tables(0).Rows(i)("T02NextPaidday")).ToOADate
                strTotal = strTotal + Val(recT03.Tables(0).Rows(i)("T07Amount"))
                strAdvance = strAdvance + Val(recT03.Tables(0).Rows(i)("T07Amount"))

                ' pbCount.Value = pbCount.Value + 1
                i = i + 1
            Next
            '----------------------------------------------------------------------------------------------
            PrintLine(1, TAB(2), "", TAB(18), "", TAB(42), "", TAB(59), "", TAB(68), "------------", TAB(80), "--------------------------")
            PrintLine(1, TAB(2), "", TAB(18), "", TAB(42), "", TAB(59), "", TAB(72), VB6.Format(nUserTotal, "#.00"), TAB(86), VB6.Format(strAdvance, "#.00"), TAB(97), VB6.Format(strTotal, "#.00"))
            'PrintLine(1, "                                                                                                                            ----------------")
            ' PrintLine(1, "                                                                                                                          ", Microsoft.VisualBasic.Right(Space(2) & VB6.Format(nUserTotal, "#.00"), 18))
            'lblPro.Text = "Complete .........."
            FileClose(1)
            Display("Lmxprint.rpt")
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                FileClose(1)
            End If
        End Try
    End Function
    Function Display(ByVal a_strFileName As String)
        Dim SysInfoPath As String
        SysInfoPath = "C:\Program Files\windows NT\Accessories"
        If (Dir(SysInfoPath & "\wordpad.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\wordpad.EXE " & a_strFileName
            Call Shell(SysInfoPath, AppWinStyle.MaximizedFocus)
        End If
    End Function
    Sub PageHeader()
        'Chr (27) & Chr(67) & Chr(96) & Chr(27) & Chr(77)
        PrintLine(1, "Run Date   : " & VB6.Format(Now, "dd/mm/yyyy"))
        PrintLine(1, "From Date  : " & VB6.Format(txtDate.Text, "dd/mm/yyyy"))
        PrintLine(1, "To Date    : " & VB6.Format(txtTo.Text, "dd/mm/yyyy"))
        PrintLine(1, Space(23) & "                 Ishara Jewelers - Main Street- Kegalla")
        ' PrintLine(1, "Receiving Officer : " & dsUser.Tables(0).Rows(Y)("T04User"))
        PrintLine(1, "Sales Report")
        'PrintLine(1, "Area Name : " & txtRoot.Text)
        PrintLine(1, "---------------------------------------------------------------------------------------------------------------")
        PrintLine(1, "  Invoice No     Item Name                Wight(grm)       Price      Discount      Advance     Total   ")
        PrintLine(1, "---------------------------------------------------------------------------------------------------------------")
        PrintLine(1, "")

    End Sub
End Class