
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmQualityReport
    Dim Clicked As String
    Dim _FualtCode As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' cboDep.ToggleDropdown()

        txtDate.Text = Today
        txtTo.Text = Today
        cmdEdit.Enabled = True
    End Sub
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        'cmdSave.Enabled = False
        cmdAdd.Focus()
    End Sub


    Private Sub frmQualityReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDate.Text = Today
        txtTo.Text = Today

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m03 As DataSet
        Dim Sql As String

        Try
            'Load Production Order No

            Sql = "select M02Name as [Fault Name] from M02Fault"
            m03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboFrom
                .DataSource = m03
                .Rows.Band.Columns(0).Width = 340
            End With


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Sub


    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String


        'Dim con = New SqlConnection()
        'con = DBEngin.GetConnection()
        Dim recGRNheader As DataSet
        Dim recStockBalance As DataSet
        ' Dim A As String
        '  Dim B As New ReportDocument

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim _FromTime As String
        Dim _ToTime As String
        Dim T01 As DataSet
        Dim nvcFieldList As String
        Dim i As Integer
        Dim _Shift As String
        Dim n_Usableqty As Double
        Dim T02 As DataSet

        Try
            StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

            nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            If chk1.Checked = True Then
                _FromTime = txtDate.Text & " " & txtTime1.Text
                _ToTime = txtTo.Text & " " & txtToTime.Text

                If Search_FualtCode() = True Then
                    nvcFieldList = "select M03MCNo,count(M03MCNo) as MCNo,T01OrderNo from T01Transaction_Header inner join T02Trans_Fault on T01RefNo=T02Ref inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T02FualtCode='" & _FualtCode & "' group by M03MCNo,T01OrderNo"
                    T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    i = 0
                    For Each DTRow1 As DataRow In T01.Tables(0).Rows

                        nvcFieldList = "select sum(T0Rollweight)as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('QP','P') group by T01OrderNo"
                        T02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                        If isValidDataset(T02) Then
                            n_Usableqty = T02.Tables(0).Rows(0)("T0Rollweight")
                            '  txtB_Qty.Text = M01.Tables(0).Rows(0)("T0Rollweight")
                        End If
                        '--------------------------------------------------------------------------------------------

                        nvcFieldList = "select sum(T0Rollweight)as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('Q','RP') group by T01OrderNo"
                        T02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                        If isValidDataset(T02) Then
                            n_Usableqty = n_Usableqty + Val(T02.Tables(0).Rows(0)("T0Rollweight"))
                            '  txtB_Qty.Text = M01.Tables(0).Rows(0)("T0Rollweight")
                        End If

                        nvcFieldList = "Insert Into R01Report(R01OrderNo,R01CutWhight,R01WorkStation,R01R_Whight)" & _
                                                               " values('" & T01.Tables(0).Rows(i)("M03MCNo") & "', '" & T01.Tables(0).Rows(i)("McNo") & "','" & netCard & "'," & Microsoft.VisualBasic.Format(((T01.Tables(0).Rows(i)("McNo") / n_Usableqty) * 25), "#.000") & ")"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)

                        'nvcFieldList = "Insert Into R01Report(R01OrderNo,R01CutWhight,R01WorkStation)" & _
                        '                                       " values('" & Trim(T01.Tables(0).Rows(i)("M03MCNo")) & "'," & T01.Tables(0).Rows(i)("McNo") & ",'" & netCard & "')"
                        'ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        i = i + 1
                    Next

                    MsgBox("Report Genarating Sucessfully", MsgBoxStyle.Information, "Textured Jersey ......")
                    transaction.Commit()
                    DBEngin.CloseConnection(connection)
                    Dim X As String
                    Dim Y As String

                    X = Hour(txtTime1.Text) & ":" & Minute(txtTime1.Text)
                    Y = Hour(txtToTime.Text) & ":" & Minute(txtTime1.Text)
                    If X = "7:30" And Y = "19:30" Then
                        _Shift = "01"
                    ElseIf X = "19:30" And Y = "7:30" Then
                        _Shift = "02"
                    End If
                    If _Shift <> "" Then
                    Else
                        _Shift = "-"
                    End If
                    A = ConfigurationManager.AppSettings("ReportPath") + "\FaultGrft.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("From", txtTo.Value)
                    B.SetParameterValue("Shift", _Shift)
                    B.SetParameterValue("Fname", cboFrom.Text)
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                Else
                    MsgBox("Please enter the correct fault code", MsgBoxStyle.Information, "Textured Jersey ........")
                End If

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Function Search_FualtCode() As Boolean

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m03 As DataSet
        Dim Sql As String

        Try
            Sql = "select * from M02Fault where M02Name='" & Trim(cboFrom.Text) & "'"
            m03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(m03) Then
                Search_FualtCode = True

                _FualtCode = m03.Tables(0).Rows(0)("M02code")
            Else
                Search_FualtCode = False
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
End Class