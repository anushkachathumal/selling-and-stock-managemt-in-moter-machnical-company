Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmAccount_Statment
    Dim _Acctype As String
    Dim _Comcode As String
    Dim c_dataCustomer1 As DataTable

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_AccStatment
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 150
            .DisplayLayout.Bands(0).Columns(2).Width = 80

            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(1).Columns(1).Width = 130
            '.DisplayLayout.Bands(1).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(2).Columns(1).Width = 90
            '.DisplayLayout.Bands(2).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(3).Columns(1).Width = 90
            '.DisplayLayout.Bands(3).Columns(1).AutoEdit = False

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Account_name()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select M01Acc_Name as [Account Name] from M01Account_Master "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboName
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 230
            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub frmAccount_Statment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("ComCode")
        Call Load_Account_name()
        Call Load_Gride2()
        txtFrom.Text = Today
        txtTo.Text = Today
        txtAcc_Limit.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

    End Sub

    Private Sub cmdExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Search_Records() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim M01 As DataSet

        Dim Value As Double

        Try
            SQL = "select * from M01Account_Master where  M01Acc_Code='" & Trim(txtAcc_No.Text) & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                cboStatus.Text = T01.Tables(0).Rows(0)("M01Status")
                _Acctype = Trim(T01.Tables(0).Rows(0)("M01Acc_Type"))
                cboName.Text = T01.Tables(0).Rows(0)("M01Acc_Name")
                'If Not IsDBNull(T01.Tables(0).Rows(0)("M01Address")).Equals = True Then
                txtAddress.Text = T01.Tables(0).Rows(0)("M01Address")
                'End If
                Value = T01.Tables(0).Rows(0)("M01Acc_Limit")
                txtAcc_Limit.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAcc_Limit.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtTp.Text = T01.Tables(0).Rows(0)("M01TP")
                txtFrom.Text = T01.Tables(0).Rows(0)("M01DOC")
                txtBuild.Text = T01.Tables(0).Rows(0)("M01DOC")

                txtStatus.Text = "M/S"

                SQL = "select * from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Acc_Code='" & _Comcode & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(M01) Then
                    txtCus.Text = M01.Tables(0).Rows(0)("M01Acc_Name")
                End If
                'txtTp.Text = T01.Tables(0).Rows(0)("M01TP")
                Search_Records = True
                Call Load_Gride_Withdata()

                'Call Load_Gridewith_Detail()
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_Records1() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim M01 As DataSet

        Dim Value As Double

        Try
            SQL = "select * from M01Account_Master where  M01Acc_Name='" & Trim(cboName.Text) & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                cboStatus.Text = T01.Tables(0).Rows(0)("M01Status")
                _Acctype = Trim(T01.Tables(0).Rows(0)("M01Acc_Type"))
                txtAcc_No.Text = T01.Tables(0).Rows(0)("M01Acc_Code")
                txtAddress.Text = T01.Tables(0).Rows(0)("M01Address")
                'Value = T01.Tables(0).Rows(0)("M01Acc_Limit")
                'txtAcc_Limit.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtAcc_Limit.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtTp.Text = T01.Tables(0).Rows(0)("M01TP")
                txtFrom.Text = T01.Tables(0).Rows(0)("M01DOC")
                txtBuild.Text = T01.Tables(0).Rows(0)("M01DOC")

                txtStatus.Text = "M/S"

                SQL = "select * from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Acc_Code='" & _Comcode & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(M01) Then
                    txtCus.Text = M01.Tables(0).Rows(0)("M01Acc_Name")
                End If
                'txtTp.Text = T01.Tables(0).Rows(0)("M01TP")

                Call Load_Gride_Withdata()

                'Call Load_Gridewith_Detail()
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub cboName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboName.AfterCloseUp
        Call Search_Records1()
    End Sub

    Private Sub cboName_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboName.InitializeLayout

    End Sub

    Private Sub txtAcc_No_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAcc_No.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()

        End If
    End Sub

    Function Load_Gride_Withdata()
        Dim sql As String
        Dim i As Integer
        Dim M01 As DataSet
        Dim _St As String
        Dim Value As Double
        Dim _OB As Double
        Dim M02 As DataSet

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            Call Load_Gride2()
            _OB = 0

            sql = "SELECT M01OB_Chq,M01Acc_Limit FROM M01Account_Master where  M01DOC<='" & txtFrom.Text & "' and M01Acc_Code='" & Trim(txtAcc_No.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, sql)
            If isValidDataset(M01) Then
                _OB = Val(M01.Tables(0).Rows(0)("M01Acc_Limit"))
            End If

            sql = "select sum(T05Credit) as T05Credit,sum(T05Debit) as T05Debit from T05Acc_Trans where  T05Date<'" & txtFrom.Text & "' and T05Acc_No='" & Trim(txtAcc_No.Text) & "' group by T05Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, sql)
            If isValidDataset(M01) Then
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Date") = Microsoft.VisualBasic.Day(txtFrom.Text) & "/" & Month(txtFrom.Text) & Year(txtFrom.Text)
                newRow("Description") = "Account Balance B/F"
                newRow("Ref.No") = "N/A"
                Value = M01.Tables(0).Rows(0)("T05Credit") + _OB
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Credit") = _St
                Value = M01.Tables(0).Rows(0)("T05Debit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Debit") = _St



                c_dataCustomer1.Rows.Add(newRow)
            Else
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Date") = Microsoft.VisualBasic.Day(txtFrom.Text) & "/" & Month(txtFrom.Text) & "/" & Year(txtFrom.Text)
                newRow("Description") = "Account Balance B/F"
                newRow("Ref.No") = "N/A"


                Value = _OB
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Debit") = _St
                newRow("Credit") = "0.00"
                c_dataCustomer1.Rows.Add(newRow)
            End If

            i = 0
            sql = "select * from T05Acc_Trans where  T05Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and T05Acc_No='" & Trim(txtAcc_No.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, sql)
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Date") = Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T05Date")) & "/" & Month(M01.Tables(0).Rows(i)("T05Date")) & "/" & Year(M01.Tables(0).Rows(i)("T05Date"))
                newRow("Description") = Trim(M01.Tables(0).Rows(i)("T05Remark"))
                newRow("Ref.No") = Trim(M01.Tables(0).Rows(i)("T05Ref_No"))
                Value = M01.Tables(0).Rows(i)("T05Credit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Credit") = _St
                Value = M01.Tables(0).Rows(i)("T05Debit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Debit") = _St

                sql = "select * from T04Chq_Trans where T04Status='RP' and T04Ref_No='" & Trim(M01.Tables(0).Rows(i)("T05Ref_No")) & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, sql)
                If isValidDataset(M02) Then
                    newRow("Due Date") = Microsoft.VisualBasic.Day(M02.Tables(0).Rows(0)("T04DOR")) & "/" & Month(M02.Tables(0).Rows(0)("T04DOR")) & "/" & Year(M02.Tables(0).Rows(0)("T04DOR"))
                    newRow("Bank Name") = M02.Tables(0).Rows(0)("T04Bank_Name")

                End If
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next

            Dim _Credit As Double
            Dim _Debit As Double

            _Credit = 0
            _Debit = 0
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                _Credit = _Credit + CDbl(UltraGrid1.Rows(i).Cells(3).Value)
                _Debit = _Debit + CDbl(UltraGrid1.Rows(i).Cells(4).Value)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Credit
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Credit") = _St
            Value = _Debit
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Debit") = _St
            c_dataCustomer1.Rows.Add(newRow1)


            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            newRow1("Date") = ""
            c_dataCustomer1.Rows.Add(newRow2)

            Dim newRow3 As DataRow = c_dataCustomer1.NewRow
            If _Debit < _Credit Then '
                newRow3("Description") = "Account Balance"
                Value = _Debit - _Credit
                If Value < 0 Then
                    Value = -Value
                End If
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow3("Credit") = _St
            Else
                newRow3("Description") = "Account Balance"
                Value = _Debit - _Credit
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow3("Debit") = _St
            End If
            c_dataCustomer1.Rows.Add(newRow3)

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim i As Integer
        Dim result1 As String
        Dim Index As Integer
        Dim M01 As DataSet
        Dim _Credit As Double
        Dim _Debit As Double
        Dim _OB As Double

        Try
            _Credit = 0
            _Debit = 0

            _OB = 0

            nvcFieldList1 = "SELECT M01OB_Chq,M01Acc_Limit FROM M01Account_Master where  M01DOC<='" & txtFrom.Text & "' and M01Acc_Code='" & Trim(txtAcc_No.Text) & "' "
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _OB = M01.Tables(0).Rows(0)("M01Acc_Limit")
            End If
            nvcFieldList1 = "delete from R01Report where  R01User='" & strDisname & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Index = 1

            nvcFieldList1 = "select sum(T05Credit) as T05Credit,sum(T05Debit) as T05Debit from T05Acc_Trans where T05Date<'" & txtFrom.Text & "' and T05Acc_No='" & Trim(txtAcc_No.Text) & "' group by T05Com_Code"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                nvcFieldList1 = "Insert Into R01Report(R01Date,R01Remark,R01Ref,R01Credit,R01Debit,R01Com,R01User,R01Index)" & _
                                                              " values('" & (Trim(txtFrom.Text)) & "', 'Account Balance B/F','N/A','" & M01.Tables(0).Rows(0)("Credit") & "','" & M01.Tables(0).Rows(0)("Debit") + _OB & "','" & _Comcode & "','" & strDisname & "'," & Index & ")"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                _Credit = M01.Tables(0).Rows(0)("T05Credit") + _OB
                _Debit = M01.Tables(0).Rows(0)("T05Debit")
            Else
                nvcFieldList1 = "Insert Into R01Report(R01Date,R01Remark,R01Ref,R01Credit,R01Debit,R01Com,R01User,R01Index)" & _
                                                         " values('" & (Trim(txtFrom.Text)) & "', 'Account Balance B/F','N/A','0','" & _OB & "','" & _Comcode & "','" & strDisname & "'," & Index & ")"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            Index = Index + 1
            i = 0
            nvcFieldList1 = "select * from T05Acc_Trans where  T05Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and T05Acc_No='" & Trim(txtAcc_No.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                nvcFieldList1 = "Insert Into R01Report(R01Date,R01Remark,R01Ref,R01Credit,R01Debit,R01Com,R01User,R01Index)" & _
                                                                        " values('" & M01.Tables(0).Rows(i)("T05Date") & "','" & Trim(M01.Tables(0).Rows(i)("T05Remark")) & "','" & Trim(M01.Tables(0).Rows(i)("T05Ref_No")) & "','" & M01.Tables(0).Rows(i)("T05Credit") & "','" & M01.Tables(0).Rows(i)("T05Debit") & "','" & _Comcode & "','" & strDisname & "'," & Index & ")"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                _Credit = _Credit + M01.Tables(0).Rows(i)("T05Credit")
                _Debit = _Debit + M01.Tables(0).Rows(i)("T05Debit")
                Index = Index + 1
                i = i + 1
            Next

            transaction.Commit()



            '_Credit = 0
            '_Debit = 0
            'i = 0
            'For Each uRow As UltraGridRow In UltraGrid1.Rows
            '    _Credit = _Credit + CDbl(UltraGrid1.Rows(i).Cells(3).Value)
            '    _Debit = _Debit + CDbl(UltraGrid1.Rows(i).Cells(4).Value)

            '    i = i + 1
            'Next

            Dim B As New ReportDocument
            Dim A As String
            Dim Value As Double
            Dim _St As String

            A = ConfigurationManager.AppSettings("ReportPath") + "\AccountStatment.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            B.SetParameterValue("Account", cboName.Text)
            B.SetParameterValue("From", txtFrom.Text)
            B.SetParameterValue("To", txtTo.Text)

            If _Credit > (_Debit + _OB) Then
                Value = _Credit - (_Debit + _OB)
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                B.SetParameterValue("Credit", _St)
                B.SetParameterValue("Debit", "-")
            Else

                Value = (_Debit + _OB) - _Credit
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                B.SetParameterValue("Credit", " ")
                B.SetParameterValue("Debit", _St)
            End If
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Ref_No}=" & txtEntry.Text & " and {T01Transaction_Header.T01Trans_Type}='GR' and {T01Transaction_Header.T01Com_Code}='" & _Comcode & "'"
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(connection)
                'connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub txtFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFrom.TextChanged
        If IsDate(txtFrom.Text) Then
            If cboName.Text <> "" Then
                Call Load_Gride_Withdata()
            End If
        End If
    End Sub

    Private Sub txtFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFrom.ValueChanged

    End Sub

    Private Sub txtAcc_No_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAcc_No.ValueChanged

    End Sub
End Class