Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Public Class frmDailySummery
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim nvcFieldList As String
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        OPR0.Enabled = True
        'OPR3.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        cmdSave.Enabled = True

        Try
            nvcFieldList = "Delete from T23Report"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            transaction.Commit()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        OPR0.Enabled = False
        'OPR3.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Dim dsUser As DataSet
        Dim ncQryType As String
        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVccode As String
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim SQL As String
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim DSUSER1 As DataSet

        ' Dim con = New SqlConnection()
        'con = DBEngin.GetConnection()
        Try
            Dim strItemtot As Double
            Dim X As Integer
            Dim STRTOTAL As Double
            Dim strDis As String
            STRTOTAL = 0
            i = 0
            X = 0
            strDis = "Motocycle Sales "
            nvcFieldList = "Insert Into T23Report(T23Description,T23Status)" & _
                                                                                 " values('" & strDis & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            SQL = "select * from T01MotocycleSales where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' AND T01Status='A'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow As DataRow In dsUser.Tables(0).Rows
              
                SQL = "SELECT * FROM T02PaymentHeader WHERE T02InvoiceNo='" & dsUser.Tables(0).Rows(i)("T01InvoiceNo") & "' AND T02Status='A'"
                DSUSER1 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(DSUSER1) Then
                    STRTOTAL = Val(STRTOTAL) + Val(DSUSER1.Tables(0).Rows(0)("T02Downpayment"))
                    strDis = "Motocycle Sales - Credit Advance/ " + CStr(dsUser.Tables(0).Rows(i)("T01InvoiceNo"))
                    nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status)" & _
                                                                                         " values('" & strDis & "','" & DSUSER1.Tables(0).Rows(0)("T02Downpayment") & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
                Else
                    STRTOTAL = Val(STRTOTAL) + Val(dsUser.Tables(0).Rows(i)("T01Total"))
                    strDis = "Motocycle Sales - Cash/ " + CStr(dsUser.Tables(0).Rows(i)("T01InvoiceNo"))
                    nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status)" & _
                                                                                         " values('" & strDis & "','" & dsUser.Tables(0).Rows(i)("T01Total") & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
                End If

                i = i + 1
            Next

            'STRTOTAL = Val(STRTOTAL) + Val(dsUser.Tables(0).Rows(i)("T01Total"))
            strDis = "Total - Motocycle Sales"
            nvcFieldList = "Insert Into T23Report(T23Description,T23Amount1,T23Status)" & _
                                                                                 " values('" & strDis & "','" & STRTOTAL & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                " values('A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            SQL = "select sum(T03Amount)as total from T03PaymentFluter where T03Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T03Status='A'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                If IsDBNull(dsUser.Tables(0).Rows(0)("total")) Then

                Else
                    STRTOTAL = Val(STRTOTAL) + Val(dsUser.Tables(0).Rows(0)("total"))
                    strDis = "Credit Collection "
                    nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status,T23Amount1)" & _
                                                                                         " values('" & strDis & "','" & dsUser.Tables(0).Rows(0)("Total") & "','A','" & dsUser.Tables(0).Rows(0)("Total") & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
                End If
            End If

                '------------------------------------------------------------------------->>
                nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                    " values('A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

            SQL = "select * FROM T04ItemSalesHeader WHERE T04Date BETWEEN '" & txtDate.Text & "' AND '" & txtTo.Text & "' AND T04Status='A'"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    nvcFieldList = "Insert Into T23Report(T23Status,T23Description)" & _
                                                      " values('A','Other Item Sales')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                strItemtot = 0
                i = 0
                SQL = "select max(M02ItemName) as MName,sum(T21Total) as Total from T21Itemsales_Fluter inner join  M02CreateItem on M02CreateItem.M02ItemCode=T21Itemsales_Fluter.T21ItemCode inner join T04ItemSalesHeader on T04ItemSalesHeader.T04InvoiceNo=T21Itemsales_Fluter.T21InvoiceNo where T04Status='A' and T04Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T21ItemCode"
                    DSUSER1 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                For Each DTRow As DataRow In DSUSER1.Tables(0).Rows
                    If IsDBNull(DSUSER1.Tables(0).Rows(i)("Total")) Then
                    Else
                        strItemtot = Val(strItemtot) + Val(Val(DSUSER1.Tables(0).Rows(i)("Total")))
                        STRTOTAL = Val(STRTOTAL) + Val(DSUSER1.Tables(0).Rows(i)("Total"))
                        strDis = DSUSER1.Tables(0).Rows(i)("MName")
                        nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status)" & _
                                                                                             " values('" & strDis & "','" & DSUSER1.Tables(0).Rows(i)("Total") & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    End If
                    i = i + 1
                Next
                End If

                strDis = "Total Sales Spairpart"
                nvcFieldList = "Insert Into T23Report(T23Description,T23Amount1,T23Status)" & _
                                                                                     " values('" & strDis & "','" & strItemtot & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                    " values('A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)
                '-------------------------------------------------------------------------->>
                'OTHER INCOME
                SQL = "select sum(T18Amount) as total from T18OtherIncome where T18Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T18Status='A'"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                strDis = "Other Income"
                If IsDBNull(dsUser.Tables(0).Rows(0)("total")) Then
                Else
                    STRTOTAL = Val(STRTOTAL) + Val(dsUser.Tables(0).Rows(0)("total"))
                    nvcFieldList = "Insert Into T23Report(T23Description,T23Amount1,T23Status,T23Amount)" & _
                                                                                         " values('" & strDis & "','" & dsUser.Tables(0).Rows(0)("total") & "','A','" & dsUser.Tables(0).Rows(0)("total") & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                        " values('A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
                End If
            End If
            '------------------------------------------------------------------------------>>>
            'INSURANCE COMMITION
            i = 0
            strItemtot = 0
            SQL = "SELECT SUM(T17Amount) AS TOTAL,MAX(M08Description) AS NAME FROM T17Income_Insucommision INNER JOIN M08insuranceCompany ON M08insuranceCompany.M08CompanyCode=T17Income_Insucommision.T17Companycode WHERE T17Date BETWEEN '" & txtDate.Text & "' AND '" & txtTo.Text & "' AND T17Status='A' AND T17PaymentMethod='001' GROUP BY T17Companycode"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                nvcFieldList = "Insert Into T23Report(T23Status,T23Description)" & _
                                                 " values('A','Income - Insurance Commision')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    If IsDBNull(dsUser.Tables(0).Rows(i)("Total")) Then
                    Else
                        strItemtot = Val(strItemtot) + Val(Val(dsUser.Tables(0).Rows(i)("Total")))
                        STRTOTAL = Val(STRTOTAL) + Val(dsUser.Tables(0).Rows(i)("Total"))
                        strDis = dsUser.Tables(0).Rows(i)("Name")
                        nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status)" & _
                                                                                             " values('" & strDis & "','" & dsUser.Tables(0).Rows(i)("Total") & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    End If
                    i = i + 1
                Next
            End If

            strDis = "Total Insurance Commition"
            nvcFieldList = "Insert Into T23Report(T23Description,T23Amount1,T23Status)" & _
                                                                                 " values('" & strDis & "','" & strItemtot & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                " values('A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '----------------------------------------------------------------------------------------->>
            i = 0
            strItemtot = 0
            SQL = "select sum(T16Amount) as total,max(M08Description) as name from T16Payment_Insurance inner join M08insuranceCompany on M08insuranceCompany.M08CompanyCode=T16Payment_Insurance.T16Companycode where T16Status='A' and T16PaymentMethod='001' and T16Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T16Companycode"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                nvcFieldList = "Insert Into T23Report(T23Status,T23Description)" & _
                                                 " values('A','Payment - Insurance Commision')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    If IsDBNull(dsUser.Tables(0).Rows(i)("Total")) Then
                    Else
                        strItemtot = Val(strItemtot) + Val(Val(dsUser.Tables(0).Rows(i)("Total")))
                        STRTOTAL = Val(STRTOTAL) - Val(dsUser.Tables(0).Rows(i)("Total"))
                        strDis = dsUser.Tables(0).Rows(i)("Name")
                        nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status)" & _
                                                                                             " values('" & strDis & "','" & -(dsUser.Tables(0).Rows(i)("Total")) & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    End If
                    i = i + 1
                Next
            End If

            strDis = "Total Insurance Payment"
            nvcFieldList = "Insert Into T23Report(T23Description,T23Amount1,T23Status)" & _
                                                                                 " values('" & strDis & "','" & -(strItemtot) & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                " values('A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '---------------------------------------------------------------------------------->>
            'PAYMENT FOR COLLECTOR

            i = 0
            strItemtot = 0
            SQL = "select sum(T20Amount) as total,max(M12Name) as name from T20Doccument_Account inner join M12DocumentCollector on M12DocumentCollector.M12Code=T20Doccument_Account.T20CollectorCode where T20Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T20CollectorCode"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                nvcFieldList = "Insert Into T23Report(T23Status,T23Description)" & _
                                                 " values('A','Payment - Document Collector')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    If IsDBNull(dsUser.Tables(0).Rows(i)("Total")) Then
                    Else
                        strItemtot = Val(strItemtot) + Val(Val(dsUser.Tables(0).Rows(i)("Total")))
                        STRTOTAL = Val(STRTOTAL) - Val(dsUser.Tables(0).Rows(i)("Total"))
                        strDis = dsUser.Tables(0).Rows(i)("Name")
                        nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status)" & _
                                                                                             " values('" & strDis & "','" & -(dsUser.Tables(0).Rows(i)("Total")) & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    End If
                    i = i + 1
                Next
            End If

            strDis = "Total Payment Document Collector"
            nvcFieldList = "Insert Into T23Report(T23Description,T23Amount1,T23Status)" & _
                                                                                 " values('" & strDis & "','" & -(strItemtot) & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                " values('A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '-------------------------------------------------------------------->>>>>
            'PAYMENT FOR OTHER EXPENCES

            i = 0
            strItemtot = 0
            SQL = "select sum(T12Amount) as total,max(M11Description) as name from T12LegerAccountEntery inner join M11LegerAccount on M11LegerAccount.M11AccCode=T12LegerAccountEntery.T12Acccode where T12Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T12Acccode"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                nvcFieldList = "Insert Into T23Report(T23Status,T23Description)" & _
                                                 " values('A','Payment - Expences')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    If IsDBNull(dsUser.Tables(0).Rows(i)("Total")) Then
                    Else
                        strItemtot = Val(strItemtot) + Val(Val(dsUser.Tables(0).Rows(i)("Total")))
                        STRTOTAL = Val(STRTOTAL) - Val(dsUser.Tables(0).Rows(i)("Total"))
                        strDis = dsUser.Tables(0).Rows(i)("Name")
                        nvcFieldList = "Insert Into T23Report(T23Dis2,T23Amount,T23Status)" & _
                                                                                             " values('" & strDis & "','" & -(dsUser.Tables(0).Rows(i)("Total")) & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    End If
                    i = i + 1
                Next
            End If

            strDis = "Total Payment for Expences"
            nvcFieldList = "Insert Into T23Report(T23Description,T23Amount1,T23Status)" & _
                                                                                 " values('" & strDis & "','" & -(strItemtot) & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            nvcFieldList = "Insert Into T23Report(T23Status)" & _
                                                " values('A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '------------------------------------------------------------------------------------>>
            'TOTAL BALANCE
            nvcFieldList = "Insert Into T23Report(T23Status,T23Amount1)" & _
                                               " values('A','" & STRTOTAL & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            transaction.Commit()
            A = ConfigurationManager.AppSettings("ReportPath") + "\DailySumm.rpt"


            StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

            B.Load(A.ToString)
            B.SetParameterValue("Todate", txtTo.Value)
            B.SetParameterValue("Fromdate", txtDate.Value)
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            'frmReport.CrystalReportViewer1.SelectionFormula = "{T01MotocycleSales.T01Date} in " & CDate(txtDate.Text) & " to " & CDate(txtTo.Text) & ""
            ' frmReport.CrystalReportViewer1.SelectionFormula = "{T03PaymentFluter.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
            frmReport.MdiParent = MDIMain
            'frmReport.Show()
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class