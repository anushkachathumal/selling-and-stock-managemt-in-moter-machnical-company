Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmDuplicate_Barcode
    Dim Clicked As String
    Dim c_dataCustomer As DataTable
    Dim _CPI As String

    Function Load_Order()
        'load NSL Combo box
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M03OrderNo as [Order No],max(M03Quality) as [Quality No],max(M03Material) as [Material],max(M03Yarnstock) as [Yarn Stock Code] from M03Knittingorder inner join T01Transaction_Header on T01OrderNo=M03OrderNo group by M03OrderNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboOrder
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 190
                .Rows.Band.Columns(1).Width = 90
                .Rows.Band.Columns(2).Width = 90
                .Rows.Band.Columns(3).Width = 240
                ' .Rows.Band.Columns(4).Width = 110
                '  .Rows.Band.Columns(5).Width = 110

            End With

            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        ' OPR0.Enabled = True
        'OPR1.Enabled = True
        OPR2.Enabled = True
        OPR5.Enabled = True
        cmdAdd.Enabled = False
        cmdDelete.Enabled = False
        cboOrder.ToggleDropdown()
        'txtFL.Focus()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR2, OPR5)
        Clicked = ""
        'OPR0.Enabled = False
        'OPR1.Enabled = False
        OPR2.Enabled = False
        cmdAdd.Enabled = True
        cmdDelete.Enabled = False
        'chk2.Checked = False
        'lblA.Text = "FL No"
        cmdAdd.Focus()
        ' grid_load()


       
    End Sub


    Function Find_Orderdetailes() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Find_Orderdetailes = False
            Sql = "select * from M03Knittingorder where M03OrderNo='" & Trim(cboOrder.Text) & "' and M03Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Find_Orderdetailes = True
                With M01
                    txtQuality.Text = .Tables(0).Rows(0)("M03Quality")
                    txtMaterial.Text = .Tables(0).Rows(0)("M03Material")
                    txtLine.Text = .Tables(0).Rows(0)("M03CuttingLine")
                    txtY_Code.Text = .Tables(0).Rows(0)("M03Yarnstock")
                    txtY_Type.Text = .Tables(0).Rows(0)("M03YarnType")
                    ' txtTotal.Text = .Tables(0).Rows(0)("M03Orderqty")
                    txtRoot.Text = .Tables(0).Rows(0)("M03Root")
                    txtMC.Text = .Tables(0).Rows(0)("M03MCNo")
                    txtLine.Text = .Tables(0).Rows(0)("M03LineItem")
                    Sql = "select T01RollNo as [Roll No] from T01Transaction_Header where T01OrderNo='" & Trim(cboOrder.Text) & "' group by T01RollNo"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    With cboRoll
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 190
                        ' .Rows.Band.Columns(1).Width = 90
                        '.Rows.Band.Columns(2).Width = 90
                        '.Rows.Band.Columns(3).Width = 240
                    End With

                    cmdSave.Enabled = True
                    ' cmdDelete.Enabled = True
                End With
            Else
                Find_Orderdetailes = False
                Me.txtRoot.Text = ""
                '  Me.txtTotal.Text = ""
                Me.txtQuality.Text = ""
                Me.txtLine.Text = ""
                Me.txtMaterial.Text = ""
                Me.txtY_Type.Text = ""
                Me.txtY_Code.Text = ""
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboOrder_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboOrder.AfterCloseUp
        Call Find_Orderdetailes()
    End Sub
    Private Sub cboOrder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboOrder.KeyUp
        If e.KeyCode = 13 Then
            Call Find_Orderdetailes()
            cboRoll.ToggleDropdown()
        End If
    End Sub

    Private Sub frmDuplicate_Barcode_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFdate.Text = Today
        Call Load_Order()
        'txtTotal.ReadOnly = True
        'txtQ_Qty.ReadOnly = True

        'txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtK_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtQ_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Me.Close()
    End Sub

    Private Sub cboOrder_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboOrder.InitializeLayout

    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim A As String
        Dim B As New ReportDocument
        Dim n_status As String

        Dim M02 As DataSet

        Dim i As Integer
        Dim x As Integer
        Try

        
            Sql = "select * from T01Transaction_Header where T01OrderNo='" & Trim(cboOrder.Text) & "' and T01RollNo='" & Trim(cboRoll.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Call CPI_Max(M01.Tables(0).Rows(i)("T01RefNo"))
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    n_status = Trim(M01.Tables(0).Rows(i)("T01Status"))
                    If Trim(M01.Tables(0).Rows(i)("T01Status")) = "P" Or Trim(M01.Tables(0).Rows(i)("T01Status")) = "F" Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Barcode.rpt"
                        B.Load(A.ToString)
                        B.SetDatabaseLogon("sa", "tommya")
                        B.SetParameterValue("CPI", _CPI)
                        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                        frmReport.CrystalReportViewer1.DisplayToolbar = True
                        frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01refno}=" & Trim(M01.Tables(0).Rows(i)("T01RefNo")) & " and {T01Transaction_Header.T01Status} = '" & Trim(M01.Tables(0).Rows(i)("T01Status")) & "'"
                        frmReport.Refresh()
                        ' frmReport.MdiParent = MDIMain
                        ' myReport.PrinttoPrinter(1, True, 0, 0)
                        ' frmReport.CrystalReportViewer1.PrintReport()
                        frmReport.CrystalReportViewer1.PrintReport()
                    ElseIf Trim(M01.Tables(0).Rows(i)("T01Status")) = "QP" Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Qpass.rpt"
                        B.Load(A.ToString)
                        B.SetDatabaseLogon("sa", "tommya")
                        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                        frmReport.CrystalReportViewer1.DisplayToolbar = True
                        frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01refno}=" & M01.Tables(0).Rows(i)("T01RefNo") & " and {T01Transaction_Header.T01Status} = '" & Trim(M01.Tables(0).Rows(i)("T01Status")) & "'"
                        frmReport.Refresh()
                        ' frmReport.MdiParent = MDIMain
                        ' myReport.PrinttoPrinter(1, True, 0, 0)
                        ' frmReport.CrystalReportViewer1.PrintReport()
                        frmReport.CrystalReportViewer1.PrintReport()

                    ElseIf Trim(M01.Tables(0).Rows(i)("T01Status")) = "QR" Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\QReject.rpt"
                        B.Load(A.ToString)
                        B.SetDatabaseLogon("sa", "tommya")
                        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                        frmReport.CrystalReportViewer1.DisplayToolbar = True
                        frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01refno}=" & M01.Tables(0).Rows(i)("T01RefNo") & " and {T01Transaction_Header.T01Status} = '" & Trim(M01.Tables(0).Rows(i)("T01Status")) & "'"
                        frmReport.Refresh()
                        ' frmReport.MdiParent = MDIMain
                        ' myReport.PrinttoPrinter(1, True, 0, 0)
                        ' frmReport.CrystalReportViewer1.PrintReport()
                        frmReport.CrystalReportViewer1.PrintReport()

                    ElseIf Trim(M01.Tables(0).Rows(i)("T01Status")) = "Q" Or Trim(M01.Tables(0).Rows(i)("T01Status")) = "RP" Then
                        Dim _QReasons As String
                        Dim _Status As String

                        x = 0
                        Sql = "select T01RefNo,T07F_Code from  T07Q_Reason inner join T01Transaction_Header on T01RefNo=T07RefNo where T01RefNo='" & Trim(M01.Tables(0).Rows(i)("T01RefNo")) & "' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        For Each DTRow2 As DataRow In M02.Tables(0).Rows
                            If x = 0 Then
                                _QReasons = Trim(M02.Tables(0).Rows(x)("T07F_Code"))
                            Else
                                _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(x)("T07F_Code"))
                            End If
                            _Status = Trim(M02.Tables(0).Rows(i)("T01RefNo"))
                            x = x + 1

                        Next
                        A = ConfigurationManager.AppSettings("ReportPath") + "\QBarcode.rpt"
                        'date = "4/1/2011"

                        '  StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Value).Day, "0#") & ", 00, 00, 00)"
                        '  StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                        B.Load(A.ToString)
                        B.SetDatabaseLogon("sa", "tommya")
                        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                        B.SetParameterValue("Reason", _QReasons)
                        B.SetParameterValue("CPI", _CPI)
                        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                        frmReport.CrystalReportViewer1.DisplayToolbar = True
                        ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01refno}=" & Trim(_Status) & " and {T01Transaction_Header.T01Status} in ['Q', 'RP']"
                        frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01refno}=" & M01.Tables(0).Rows(i)("T01RefNo") & " and  {T01Transaction_Header.T01Status} in ['Q', 'RP']"
                        frmReport.Refresh()
                        ' frmReport.MdiParent = MDIMain
                        ' myReport.PrinttoPrinter(1, True, 0, 0)
                        frmReport.CrystalReportViewer1.PrintReport()



                    Else
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Reject.rpt"
                        B.Load(A.ToString)
                        B.SetDatabaseLogon("sa", "tommya")
                        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                        frmReport.CrystalReportViewer1.DisplayToolbar = True
                        frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01refno}=" & M01.Tables(0).Rows(i)("T01RefNo") & " and {T01Transaction_Header.T01Status} = '" & Trim(n_status) & "'"
                        frmReport.Refresh()
                        '  frmReport.MdiParent = MDIMain
                        ' myReport.PrinttoPrinter(1, True, 0, 0)
                        frmReport.CrystalReportViewer1.PrintReport()

                    End If
                    i = i + 1
                Next
            End If



            i = 0
            Sql = "select * from T01Transaction_Header  where T01OrderNo='" & Trim(cboOrder.Text) & "' and T01RollNo='" & Trim(cboRoll.Text) & "'"

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Sql = "select T05RefNo,T05Fualcode,T05Weight,T05Reperat,T05Status from T05Scrab inner join T01Transaction_Header on T01RefNo=T05RefNo where T05RefNo='" & M01.Tables(0).Rows(0)("T01RefNo") & "'"

                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                For Each DTRow2 As DataRow In M02.Tables(0).Rows
                    If x = 0 Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Cutoff.rpt"
                    ElseIf x = 1 Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Cutoff1.rpt"
                    ElseIf x = 2 Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Cutoff2.rpt"
                    ElseIf x = 3 Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Cutoff3.rpt"
                    ElseIf x = 4 Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Cutoff4.rpt"
                    ElseIf x = 5 Then
                        A = ConfigurationManager.AppSettings("ReportPath") + "\Cutoff5.rpt"
                    End If
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    ''  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    ''   B.SetParameterValue("Reason", _QReasons)
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01refno}=" & Trim(M01.Tables(0).Rows(0)("T01RefNo")) & " and {T05Scrab.T05Status}='" & Trim(M02.Tables(0).Rows(0)("T05Status")) & "'"
                    frmReport.Refresh()
                    frmReport.MdiParent = MDIMain
                    ' myReport.PrinttoPrinter(1, True, 0, 0)
                    frmReport.CrystalReportViewer1.PrintReport()
                    '  frmReport.Show()
                    x = x + 1
                Next

            End If

            'AUTO LANE COMPLETE
            Dim _roll As Integer

            Sql = "select * from C01Workstation where C01ID='" & netCard & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _roll = M01.Tables(0).Rows(0)("C01Roll")
            End If

            Sql = "select count(T01RollNo) from T01Transaction_Header where T01OrderNo='" & Trim(cboOrder.Text) & "' group by T01RollNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            Dim X1 As Integer
            Dim _StartRoll As String
            Dim _EndRoll As String
            Dim _MidRoll As String

            '  X1 = CInt(M01.Tables(0).Rows.Count) / _roll
            X1 = 0
            ' MsgBox(Val(cboRoll.Text))
            If _roll = M01.Tables(0).Rows.Count Or _roll = Val(cboRoll.Text) Then
                _StartRoll = "001"
                If _roll = 12 Then
                    _MidRoll = "0" & CStr(_roll)
                    '_EndRoll = "0" & CStr(_roll)
                Else
                    _MidRoll = "0" & CStr(_roll - 2)
                    _EndRoll = "0" & CStr(_roll)
                End If

                X1 = 1
            ElseIf (_roll * 2) = M01.Tables(0).Rows.Count Or (_roll * 2) = Val(cboRoll.Text) Then
                _StartRoll = "0" & CStr(_roll + 1)
                ' _MidRoll = "0" & CStr(_roll - 4)
                '_EndRoll = "0" & CStr(_roll * 2)

                If _roll = 12 Then
                    _MidRoll = "0" & CStr(_roll * 2)
                    '_EndRoll = "0" & CStr(_roll)
                Else
                    _MidRoll = "0" & CStr((_roll * 2) - 2)
                    _EndRoll = "0" & CStr(_roll * 2)
                End If
                X1 = 1
            ElseIf (_roll * 3) = M01.Tables(0).Rows.Count Or (_roll * 3) = Val(cboRoll.Text) Then
                _StartRoll = "0" & CStr((_roll * 2) + 1)
                '_MidRoll = "0" & CStr((_roll * 3) - 4)
                '_EndRoll = "0" & CStr(_roll * 3)
                If _roll = 12 Then
                    _MidRoll = "0" & CStr(_roll * 3)
                    '_EndRoll = "0" & CStr(_roll)
                Else
                    _MidRoll = "0" & CStr((_roll * 3) - 2)
                    _EndRoll = "0" & CStr(_roll * 3)
                End If
                X1 = 1
            ElseIf (_roll * 4) = M01.Tables(0).Rows.Count Or (_roll * 4) = Val(cboRoll.Text) Then
                _StartRoll = "0" & CStr((_roll * 3) + 1)
                '_MidRoll = "0" & CStr((_roll * 4) - 4)
                '_EndRoll = "0" & CStr(_roll * 4)
                If _roll = 12 Then
                    _MidRoll = "0" & CStr(_roll * 4)
                    '_EndRoll = "0" & CStr(_roll)
                Else
                    _MidRoll = "0" & CStr((_roll * 4) - 2)
                    _EndRoll = "0" & CStr(_roll * 4)
                End If
                X1 = 1
            ElseIf (_roll * 5) = M01.Tables(0).Rows.Count Or (_roll * 5) = Val(cboRoll.Text) Then
                _StartRoll = "0" & CStr((_roll * 4) + 1)
                '_MidRoll = "0" & CStr((_roll * 5) - 4)
                '_EndRoll = "0" & CStr(_roll * 5)
                If _roll = 12 Then
                    _MidRoll = "0" & CStr(_roll * 5)
                    '_EndRoll = "0" & CStr(_roll)
                Else
                    _MidRoll = "0" & CStr((_roll * 5) - 2)
                    _EndRoll = "0" & CStr(_roll * 5)
                End If
                X1 = 1
            End If
            If X1 >= 1 Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Auto.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and  {T01Transaction_Header.T01Roll} in " & CInt(_StartRoll) & " to " & CInt(_MidRoll) & " and not ({T01Transaction_Header.T01Status} startswith ['Q', 'QR', 'R'])"
                '  frmReport.MdiParent = MDIMain
                ' myReport.PrinttoPrinter(1, True, 0, 0)
                frmReport.CrystalReportViewer1.PrintReport()

                If _roll = 16 Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\Auto1.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and  {T01Transaction_Header.T01Roll} in " & CInt(_MidRoll) & "' to " & CInt(_EndRoll) & " and not ({T01Transaction_Header.T01Status} startswith ['Q', 'QR', 'R'])"
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and  {T01Transaction_Header.T01Roll} in " & CInt(_MidRoll) + 1 & " to " & CInt(_EndRoll) & " and not ({T01Transaction_Header.T01Status} startswith ['Q', 'QR', 'R'])"
                    '  frmReport.MdiParent = MDIMain
                    ' myReport.PrinttoPrinter(1, True, 0, 0)
                    frmReport.CrystalReportViewer1.PrintReport()
                End If
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub


    Function CPI_Max(ByVal vMax As Integer)
        Dim _Min As Double
        Dim _Max As Double
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim m02 As DataSet
        Dim x As Integer
        Dim SQL As String


        Try
            SQL = "select * from T03CPI_Reading where T03RefNo=" & vMax & ""
            m02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(m02) Then
                If CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV2"))) >= Val(Trim(m02.Tables(0).Rows(x)("T03CPIV1"))) Then
                    _Max = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV2")))
                Else
                    _Max = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV1")))
                End If

                If _Max > CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV3"))) Then
                    _Max = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV3")))
                End If

                If _Max > CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV4"))) Then
                    _Max = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV4")))
                End If

                If _Max > CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV5"))) Then
                    _Max = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV5")))
                End If


                If CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV2"))) >= Val(Trim(m02.Tables(0).Rows(x)("T03CPIV1"))) Then
                    _Min = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV1")))
                Else
                    _Min = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV2")))
                End If

                If _Min < CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV3"))) Then
                    _Min = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV3")))
                End If

                If _Min < CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV4"))) Then
                    _Min = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV4")))
                End If

                If _Min < CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV5"))) Then
                    _Min = CInt(Trim(m02.Tables(0).Rows(x)("T03CPIV5")))
                End If
                _CPI = _Min.ToString & "/" & _Max.ToString
            Else
                _CPI = "NO"
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub OPR2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPR2.Click

    End Sub
End Class