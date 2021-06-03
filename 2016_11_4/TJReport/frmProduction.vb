Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmProduction
    Dim Clicked As String
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True

        txtDate.Text = Today
        txtTo.Text = Today
        chk0.Checked = True
        cmdEdit.Enabled = True
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        'cmdSave.Enabled = False


        cmdAdd.Focus()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub


    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim _FromTime As String
        Dim _ToTime As String
        Dim _Shift As Integer

        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String


  
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim strInvo As String
        Dim strChqvalue As Double
        Dim _DisFault As String

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim T01 As DataSet
        Dim T02 As DataSet
        Dim nvcFieldList As String

        Dim Sql As String
        Dim i As Integer


        Try
            Sql = "delete from R04Report where R04Status='" & netCard & "' "
            ExecuteNonQueryText(connection, transaction, Sql)

            If chk0.Checked = False And chk1.Checked = False And chk3.Checked = False Then
                MsgBox("Please select the shift", MsgBoxStyle.Information, "Textured Jersey ........")

            Else
                If chk0.Checked = True Then
                    _Shift = 1

                    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                    StrToDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"

                    _FromTime = txtDate.Text & " " & "07:30:00"
                    _ToTime = txtDate.Text & " " & "19:30:00"

                ElseIf chk1.Checked = True Then

                  
                    _Shift = 2

                    _FromTime = txtDate.Text & " " & "07:30 PM"
                    _ToTime = System.DateTime.FromOADate(CDate(txtDate.Text).ToOADate + 1)
                    _ToTime = _ToTime & " " & "07:30 AM"

                    txtTo.Text = _ToTime

                    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                    StrToDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"

                    '_FromTime = txtDate.Text & " " & "07:30:00"
                    '_ToTime = txtDate.Text & " " & "19:30:00"

                ElseIf chk3.Checked = True Then
                    ' _Shift = "Monthly"
                    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                    StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                    _FromTime = txtDate.Text & " " & "07:30 AM"
                    _ToTime = System.DateTime.FromOADate(CDate(txtTo.Text).ToOADate + 1)
                    _ToTime = _ToTime & " " & "07:30 AM"

                End If

                i = 0
                Dim _Rollwt As Double
                Dim _Scrap As Double
                Dim _Eff As Double

                If chk3.Checked = True Then
                    Sql = "select T04Emp,sum(T04TOTAL) as T04TOTAL from T04Summery where T04Time between '" & _FromTime & "' and '" & _ToTime & "' group by T04Emp"
                Else
                    Sql = "select * from T04Summery where T04Time between '" & _FromTime & "' and '" & _ToTime & "' and T04Shift=" & _Shift & ""
                End If

                T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                For Each DTRow1 As DataRow In T01.Tables(0).Rows

                    'If i = 9 Then
                    '    MsgBox("")
                    'End If

                    _Rollwt = 0
                    Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01InsEPF='" & Trim(T01.Tables(0).Rows(i)("T04Emp")) & "' and T01Status <>'I' group by T01InsEPF"
                    T02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(T02) Then
                        _Rollwt = T02.Tables(0).Rows(0)("T0Rollweight")
                    End If

                    '------------------------------------------------SCRAP (NEW MODIFICATION ON 2012.10.03)   REQUEST BY PRADEEP
                    _Scrap = 0
                    Sql = "select sum(T0Rollweight) as T0Rollweight,SUM(T01Reject) AS REJECT from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01InsEPF='" & Trim(T01.Tables(0).Rows(i)("T04Emp")) & "' and T01Status IN ('R','QR') group by T01InsEPF"
                    T02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(T02) Then
                        _Scrap = Val(T02.Tables(0).Rows(0)("T0Rollweight"))
                        If IsDBNull((T02.Tables(0).Rows(0)("REJECT"))) Then
                        Else
                            _Scrap = _Scrap + Val(T02.Tables(0).Rows(0)("REJECT"))
                        End If
                    End If
                    '+++++++++++++++++++++++++++++++++++++++++3M CUTOFF
                    Sql = "select sum(T05Weight) as T0Rollweight from T01Transaction_Header INNER JOIN T05Scrab ON T05RefNo=T01RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01InsEPF='" & Trim(T01.Tables(0).Rows(i)("T04Emp")) & "' and T01Status <>'I' group by T01InsEPF"
                    T02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(T02) Then
                        _Scrap = Val(T02.Tables(0).Rows(0)("T0Rollweight")) + _Scrap
                    End If

                    '-------------------------------------- LESS THAN 3M

                    Sql = "select sum(T04Weight) as T0Rollweight from T01Transaction_Header INNER JOIN T04Cutoff ON T04RefNo=T01RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01InsEPF='" & Trim(T01.Tables(0).Rows(i)("T04Emp")) & "' and T01Status <> 'I' group by T01InsEPF"
                    T02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(T02) Then
                        _Scrap = Val(T02.Tables(0).Rows(0)("T0Rollweight")) + _Scrap
                    End If

                    If _Scrap > 0 Then
                        ' MsgBox(T01.Tables(0).Rows(i)("T04TOTAL"))
                        _Eff = (_Scrap / T01.Tables(0).Rows(i)("T04TOTAL")) * 100
                        _Eff = (_Scrap / _Rollwt) * 100
                    Else
                        _Eff = 0
                    End If
                    Dim n_Eff As Double
                    Dim x_Date As Integer
                    Dim _TotalInspecTime As Integer
                    _TotalInspecTime = 0
                    ' MsgBox(T01.Tables(0).Rows(i)("T04TOTAL"))
                    'MsgBox(_Rollwt + _Scrap)


                    Sql = "select sum(T18Time) as T18Time from T18Downtime where T18User='" & Trim(T01.Tables(0).Rows(i)("T04Emp")) & "' and T18Timein between '" & _FromTime & "' and '" & _ToTime & "' group by T18User"
                    T02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(T02) Then
                        _TotalInspecTime = T02.Tables(0).Rows(0)("T18Time")
                    End If

                    Dim _TotalRoll As Integer
                    Dim _AvgTime As Double
                    Dim _TOBEIRoll As Integer

                    _AvgTime = 0
                    _TotalRoll = 0
                    _TotalRoll = T01.Tables(0).Rows(i)("T04TOTAL")

                    'CALCULATE AVG TIME
                    _AvgTime = _TotalInspecTime / _TotalRoll

                    'TO BE INSPECTED ROLL
                    If _AvgTime > 0 Then
                        _TOBEIRoll = 630 / _AvgTime
                    End If
                    If _TOBEIRoll > 0 Then
                        n_Eff = _TotalRoll / _TOBEIRoll

                        n_Eff = n_Eff * 100
                    End If
                    ' _TotalInspecTime = T01.Tables(0).Rows(i)("T04TOTAL") * 9.5
                    'x_Date = System.DateTime.DaysInMonth(Year(txtDate.Text), Month(txtDate.Text))
                    'n_Eff = _TotalInspecTime / (10.5 * 60)
                    'n_Eff = n_Eff * 100

                    'n_Eff = (10.5 * 60)

                    'n_Eff = n_Eff * x_Date
                    nvcFieldList = "Insert Into R04Report(R04Emp,R04Roll,R04wt,R04TotTime,R04Eff,R04Status,R04Scrap,R04Seff)" & _
                                                                " values('" & T01.Tables(0).Rows(i)("T04Emp") & "', " & T01.Tables(0).Rows(i)("T04TOTAL") & ",'" & _Rollwt & "','" & _TotalInspecTime & "','" & n_Eff & "','" & netCard & "'," & _Scrap & "," & VB6.Format(_Eff, "#.00") & ")"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    i = i + 1
                Next

                MsgBox("Report Genarating successfully", MsgBoxStyle.Information, "Textured Jersey ..........")
                transaction.Commit()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""

                A = ConfigurationManager.AppSettings("ReportPath") + "\ProductionMonitering.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                If chk3.Checked = True Then
                    B.SetParameterValue("Shift", "Monthly")
                Else
                    B.SetParameterValue("Shift", _Shift)
                End If
                B.SetParameterValue("From", txtDate.Value)
                ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime" & P01 & ""
                frmReport.CrystalReportViewer1.SelectionFormula = "{R04Report.R04Status}='" & netCard & "'"
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
                End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(i)
            End If
        End Try
    End Sub

    Private Sub chk1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk1.CheckedChanged
        If chk1.Checked = True Then
            chk0.Checked = False
            chk3.Checked = False
        End If
    End Sub

    Private Sub chk0_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk0.CheckedChanged
        If chk0.Checked = True Then
            chk1.Checked = False
            chk3.Checked = False
        End If
    End Sub

    Private Sub chk3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk3.CheckedChanged
        If chk3.Checked = True Then
            chk1.Checked = False
            chk0.Checked = False
        End If
    End Sub
End Class