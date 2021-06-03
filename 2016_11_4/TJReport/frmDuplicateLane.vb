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
Public Class frmDuplicateLane
    Dim Clicked As String
    Dim c_dataCustomer As DataTable
    Dim _CPI As String

    Function Load_Order()
        'load NSL Combo box
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select T08OrderNo as [Order No] from T08LaneComplete group by T08OrderNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboOrder
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 190
                '  .Rows.Band.Columns(1).Width = 90
                ' .Rows.Band.Columns(2).Width = 90
                '.Rows.Band.Columns(3).Width = 240
                ' .Rows.Band.Columns(4).Width = 110
                '  .Rows.Band.Columns(5).Width = 110

            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
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
        ' OPR5.Enabled = True
        cmdAdd.Enabled = False
        cmdDelete.Enabled = False
        cboOrder.ToggleDropdown()
        Call Load_Gride()
        'txtFL.Focus()
    End Sub

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select T08BarcodeNo as [BarCode],T08Date as [Date],T08RollNo as [Roll No],T08Qty as [Qty],T08Status AS [Status] from T08LaneComplete where T08OrderNo='" & Trim(cboOrder.Text) & "' and T08Lane='" & cboRoll.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 170
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 120
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 90
                .DisplayLayout.Bands(0).Columns(3).Width = 110
                '.DisplayLayout.Bands(0).Columns(3).Width = 160
                '.DisplayLayout.Bands(0).Columns(4).Width = 100
                '.DisplayLayout.Bands(0).Columns(5).Width = 120
                '.DisplayLayout.Bands(0).Columns(6).Width = 100
            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Lane()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select T08Lane as [Lane No] from T08LaneComplete where T08OrderNo='" & cboOrder.Text & "' group by T08Lane"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRoll
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 190
                '  .Rows.Band.Columns(1).Width = 90
                ' .Rows.Band.Columns(2).Width = 90
                '.Rows.Band.Columns(3).Width = 240
                ' .Rows.Band.Columns(4).Width = 110
                '  .Rows.Band.Columns(5).Width = 110

            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
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

    Private Sub cboOrder_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboOrder.AfterCloseUp
        Call Load_Lane()

    End Sub

    Private Sub cboOrder_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboOrder.InitializeLayout

    End Sub

    Private Sub cboOrder_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboOrder.TextChanged
        Call Load_Lane()
    End Sub

    Private Sub cboRoll_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoll.AfterCloseUp
        Call Load_Gride()
        cmdSave.Enabled = True
    End Sub

    Private Sub cboRoll_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboRoll.InitializeLayout
        Call Load_Gride()
        cmdSave.Enabled = True
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR2)
        Clicked = ""
        'OPR0.Enabled = False
        'OPR1.Enabled = False
        OPR2.Enabled = False
        cmdAdd.Enabled = True
        cmdDelete.Enabled = False
        Call Load_Gride()
        cmdAdd.Focus()
        ' grid_load()
    End Sub

    Private Sub frmDuplicateLane_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        Call Load_Order()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim _StartRoll As String
        Dim _EndRoll As String
        Dim _MidRoll As String
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)

        Dim recGRNheader As DataSet
        Dim recStockBalance As DataSet
        ' Dim A As String
        Dim rsLane As DataSet
        Dim M01 As DataSet
        Dim T08 As DataSet
        Dim B As New ReportDocument
        Dim A As String

        Dim _MCNo As String
        Dim _Date As Date

        Try
            Sql = "select M03MCNo from M03Knittingorder where M03OrderNo='" & Trim(cboOrder.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _MCNo = M01.Tables(0).Rows(0)("M03MCNo")
            End If
            ''--------------------------------------------------------->>
            ''End Roll No
            'Sql = "select T01Time from T01Transaction_Header where T01OrderNo='" & Trim(cboOrder.Text) & "' and T01RollNo='" & _EndRoll & "'"
            'M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            'If isValidDataset(M01) Then
            '    _Date = M01.Tables(0).Rows(0)("T01Time")
            'End If
            Dim _XRoll As Integer


            Sql = "select * from T08LaneComplete where T08OrderNo='" & Trim(cboOrder.Text) & "' and T08Lane=" & cboRoll.Text & ""
            T08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(T08) Then
                _XRoll = T08.Tables(0).Rows.Count
            End If
            Dim _roll As Integer

            Sql = "select * from C01Workstation where C01ID='" & netCard & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _roll = M01.Tables(0).Rows(0)("C01Roll")
            End If

            Sql = "select MIN((T08RollNo)) as MinRoll ,MAX((T08RollNo)) as MaxRoll from T08LaneComplete where T08OrderNo='" & Trim(cboOrder.Text) & "' and T08Lane=" & CInt(cboRoll.Text) & " group by T08OrderNo,T08Lane"
            rsLane = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(rsLane) Then

                '    If CInt(cboRoll.Text) >= 1 And _XRoll = 12 Then
                '        A = ConfigurationManager.AppSettings("ReportPath") + "\Auto2.rpt"
                '        B.Load(A.ToString)
                '       B.SetDatabaseLogon("sa", "tommya")
                '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '        B.SetParameterValue("MC", _MCNo)
                '        B.SetParameterValue("Date", Today)
                '        B.SetParameterValue("Name", strDisname)
                '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '        frmReport.CrystalReportViewer1.DisplayToolbar = True
                '        ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                '        frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & " "
                '        frmReport.Refresh()
                '        '  frmReport.MdiParent = MDIMain
                '        ' myReport.PrinttoPrinter(1, True, 0, 0)
                '        frmReport.CrystalReportViewer1.PrintReport()

                '    ElseIf CInt(cboRoll.Text) >= 1 And _XRoll = 16 Then
                '        Dim _XFRoll As String

                '        If CInt(cboRoll.Text) = 1 Then
                '            _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                '            _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 5)
                '            _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                '            _XFRoll = "0" & (CInt(_MidRoll) + 1)
                '        ElseIf CInt(cboRoll.Text) = 2 Then
                '            _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                '            _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 5)
                '            _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                '            _XFRoll = "0" & (CInt(_MidRoll) + 1)

                '        ElseIf CInt(cboRoll.Text) = 3 Then
                '            _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                '            _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 5)
                '            _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                '            _XFRoll = "0" & (CInt(_MidRoll) + 1)
                '        ElseIf CInt(cboRoll.Text) = 4 Then
                '            _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                '            _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 5)
                '            _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                '            _XFRoll = "0" & (CInt(_MidRoll) + 1)

                '        ElseIf CInt(cboRoll.Text) = 5 Then
                '            _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                '            _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 5)
                '            _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                '            _XFRoll = "0" & (CInt(_MidRoll) + 1)
                '        End If
                '        A = ConfigurationManager.AppSettings("ReportPath") + "\Auto2.rpt"
                '        B.Load(A.ToString)
                '       B.SetDatabaseLogon("sa", "tommya")
                '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '        B.SetParameterValue("MC", _MCNo)
                '        B.SetParameterValue("Date", Today)
                '        B.SetParameterValue("Name", strDisname)
                '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '        frmReport.CrystalReportViewer1.DisplayToolbar = True
                '        ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                '        frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & " and {T08LaneComplete.T08RollNo} in '" & _StartRoll & "' to '" & _MidRoll & "' "
                '        frmReport.Refresh()
                '        '  frmReport.MdiParent = MDIMain
                '        ' myReport.PrinttoPrinter(1, True, 0, 0)
                '        frmReport.CrystalReportViewer1.PrintReport()


                '        A = ConfigurationManager.AppSettings("ReportPath") + "\Auto1.rpt"
                '        B.Load(A.ToString)
                '       B.SetDatabaseLogon("sa", "tommya")
                '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '        B.SetParameterValue("MC", _MCNo)
                '        B.SetParameterValue("Date", Today)
                '        B.SetParameterValue("Name", strDisname)
                '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '        frmReport.CrystalReportViewer1.DisplayToolbar = True
                '        ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                '        frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & " and {T08LaneComplete.T08RollNo} in '" & _XFRoll & "' to '" & _EndRoll & "' "
                '        frmReport.Refresh()
                '        '  frmReport.MdiParent = MDIMain
                '        ' myReport.PrinttoPrinter(1, True, 0, 0)
                '        frmReport.CrystalReportViewer1.PrintReport()
                '    End If
                'End If

                If CInt(cboRoll.Text) >= 1 And _roll = 9 Then

                    Dim _XFRoll As String

                    If CInt(cboRoll.Text) = 1 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 2)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)
                    ElseIf CInt(cboRoll.Text) = 2 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 2)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)

                    ElseIf CInt(cboRoll.Text) = 3 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 2)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)
                    ElseIf CInt(cboRoll.Text) = 4 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 2)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)

                    ElseIf CInt(cboRoll.Text) = 5 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 2)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)
                    End If

                    'A = ConfigurationManager.AppSettings("ReportPath") + "\Auto2.rpt"
                    'B.Load(A.ToString)
                    'B.SetDatabaseLogon("sa", "sainfinity")
                    ''  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    'B.SetParameterValue("MC", _MCNo)
                    'B.SetParameterValue("Date", Today)
                    'B.SetParameterValue("Name", strDisname)
                    'frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    'frmReport.CrystalReportViewer1.DisplayToolbar = True
                    '' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                    'frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & _Lane & " "
                    'frmReport.Refresh()
                    ''  frmReport.MdiParent = MDIMain
                    '' myReport.PrinttoPrinter(1, True, 0, 0)
                    'frmReport.CrystalReportViewer1.PrintReport()

                    A = ConfigurationManager.AppSettings("ReportPath") + "\Auto2.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    B.SetParameterValue("MC", _MCNo)
                    B.SetParameterValue("Date", Today)
                    B.SetParameterValue("Name", strDisname)
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                    'If _XRoll < 9 Then
                    '    frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & "" ' and {T08LaneComplete.T08RollNo} in '" & _StartRoll & "' to '" & _MidRoll & "' "
                    'Else
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & " and {T08LaneComplete.T08RollNo} in '" & _StartRoll & "' to '" & _MidRoll & "' "
                    ' End If
                    frmReport.Refresh()
                    '  frmReport.MdiParent = MDIMain
                    ' myReport.PrinttoPrinter(1, True, 0, 0)
                    frmReport.CrystalReportViewer1.PrintReport()

                    'If _XRoll > 9 Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\Auto1.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    B.SetParameterValue("MC", _MCNo)
                    B.SetParameterValue("Date", Today)
                    B.SetParameterValue("Name", strDisname)
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & " and {T08LaneComplete.T08RollNo} in '" & _XFRoll & "' to '" & _EndRoll & "' "
                    frmReport.Refresh()
                    '  frmReport.MdiParent = MDIMain
                    ' myReport.PrinttoPrinter(1, True, 0, 0)
                    frmReport.CrystalReportViewer1.PrintReport()
                    'End If
                ElseIf CInt(cboRoll.Text) >= 1 And _roll = 16 Then
                    Dim _XFRoll As String

                    If CInt(cboRoll.Text) = 1 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 6)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)
                    ElseIf CInt(cboRoll.Text) = 2 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 6)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)

                    ElseIf CInt(cboRoll.Text) = 3 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 6)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)
                    ElseIf CInt(cboRoll.Text) = 4 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 6)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)

                    ElseIf CInt(cboRoll.Text) = 5 Then
                        _StartRoll = Trim(rsLane.Tables(0).Rows(0)("MinRoll"))
                        _MidRoll = "0" & (CInt(Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))) - 6)
                        _EndRoll = Trim(rsLane.Tables(0).Rows(0)("MaxRoll"))
                        _XFRoll = "0" & (CInt(_MidRoll) + 1)
                    End If
                    A = ConfigurationManager.AppSettings("ReportPath") + "\Auto2.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    B.SetParameterValue("MC", _MCNo)
                    B.SetParameterValue("Date", Today)
                    B.SetParameterValue("Name", strDisname)
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & " and {T08LaneComplete.T08RollNo} in '" & _StartRoll & "' to '" & _MidRoll & "' "
                    frmReport.Refresh()
                    '  frmReport.MdiParent = MDIMain
                    ' myReport.PrinttoPrinter(1, True, 0, 0)
                    frmReport.CrystalReportViewer1.PrintReport()


                    A = ConfigurationManager.AppSettings("ReportPath") + "\Auto1.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    B.SetParameterValue("MC", _MCNo)
                    B.SetParameterValue("Date", Today)
                    B.SetParameterValue("Name", strDisname)
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & Trim(cboOrder.Text) & "' and {T01Transaction_Header.T01Status} in ('P', 'QP', 'RP') and {T01Transaction_Header.T01RollNo} in ('" & _StartRoll & "' to '" & _EndRoll & "')"
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T08LaneComplete.T08OrderNo} ='" & Trim(cboOrder.Text) & "' and  {T08LaneComplete.T08Lane}=" & CInt(cboRoll.Text) & " and {T08LaneComplete.T08RollNo} in '" & _XFRoll & "' to '" & _EndRoll & "' "
                    frmReport.Refresh()
                    '  frmReport.MdiParent = MDIMain
                    ' myReport.PrinttoPrinter(1, True, 0, 0)
                    frmReport.CrystalReportViewer1.PrintReport()
                End If

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""

            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class